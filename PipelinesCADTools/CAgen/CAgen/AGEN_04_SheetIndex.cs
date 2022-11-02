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
using System.Management;

namespace Alignment_mdi
{
    public partial class AGEN_SheetIndex : Form
    {
        private ContextMenuStrip ContextMenuStrip_xl;

        List<ObjectId> lista_del = new List<ObjectId>();

        public AGEN_SheetIndex()
        {
            InitializeComponent();
            if (Functions.is_dan_popescu() == true) checkBox_draw_automatic.Visible = true;

            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem1.Click += delete_row_Click;

            ContextMenuStrip_xl = new ContextMenuStrip();
            ContextMenuStrip_xl.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1 });

        }
        private void delete_row_Click(object sender, EventArgs e)
        {
            try
            {



                if (dataGridView_sheet_index.RowCount > 0)
                {
                    int index_grid = dataGridView_sheet_index.CurrentCell.RowIndex;
                    if (index_grid == -1)
                    {
                        return;
                    }



                    string dwg1_name = "";

                    if (dataGridView_sheet_index.Rows[index_grid].Cells[_AGEN_mainform.Col_dwg_name].Value != DBNull.Value)
                    {
                        dwg1_name = Convert.ToString(dataGridView_sheet_index.Rows[index_grid].Cells[_AGEN_mainform.Col_dwg_name].Value);
                    }

                    if (dwg1_name != "")
                    {
                        for (int j = _AGEN_mainform.dt_sheet_index.Rows.Count - 1; j >= 0; --j)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string dwg2_name = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name]);

                                if (dwg2_name.ToLower() == dwg1_name.ToLower())
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[j].Delete();
                                    label_not_saved.Visible = true;
                                    j = -1;
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

        private void dataGridView_sheet_index_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_xl.Show(Cursor.Position);
                ContextMenuStrip_xl.Visible = true;
            }
            else
            {
                ContextMenuStrip_xl.Visible = false;
            }
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_from_start_end);
            lista_butoane.Add(button_adjust_rectangle);
            lista_butoane.Add(button_delete_3d_poly);
            lista_butoane.Add(button_delete_sheet_index);
            lista_butoane.Add(button_draw_manual);
            lista_butoane.Add(button_draw_Viewport_templates);
            lista_butoane.Add(button_fill_gaps);
            lista_butoane.Add(button_insert_matchline_block_ms);
            lista_butoane.Add(button_insert_na_ms);
            lista_butoane.Add(button_place_rectangles);
            lista_butoane.Add(button_recover_matchlines);
            lista_butoane.Add(button_scan);
            lista_butoane.Add(button_open_sheet_index_xl);
            lista_butoane.Add(button_pick_middle);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_from_start_end);
            lista_butoane.Add(button_adjust_rectangle);
            lista_butoane.Add(button_delete_3d_poly);
            lista_butoane.Add(button_delete_sheet_index);
            lista_butoane.Add(button_draw_manual);
            lista_butoane.Add(button_draw_Viewport_templates);
            lista_butoane.Add(button_fill_gaps);
            lista_butoane.Add(button_insert_matchline_block_ms);
            lista_butoane.Add(button_insert_na_ms);
            lista_butoane.Add(button_place_rectangles);
            lista_butoane.Add(button_recover_matchlines);
            lista_butoane.Add(button_scan);
            lista_butoane.Add(button_open_sheet_index_xl);
            lista_butoane.Add(button_pick_middle);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public void set_dataGridView_sheet_index()
        {
            dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.EnableHeadersVisualStyles = false;
        }




        private void button_place_rectangles_Click(object sender, EventArgs e)
        {
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
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }


            if (Functions.IsNumeric(TextBox_matchline_length.Text) == false)
            {
                MessageBox.Show("no matchlines distance specified\r\nOperation aborted");
                return;
            }

            if (_AGEN_mainform.Vw_height == 0 || _AGEN_mainform.Vw_width == 0)
            {
                MessageBox.Show("you do not have the dimensions for the matchline rectangles\r\nOperation aborted");
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
                return;
            }

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you do not have picked the centerline\r\noperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_viewport_settings.Hide();
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

                _AGEN_mainform.tpage_setup.Show();
                return;
            }
            double poly_length = 0;
            Ag.WindowState = FormWindowState.Minimized;

            set_enable_false();
            Erase_viewports_templates();
            Create_ML_object_data();
            if (Functions.IsNumeric(TextBox_matchline_length.Text) == true)
            {
                _AGEN_mainform.Match_distance = Convert.ToDouble(TextBox_matchline_length.Text);
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_AGEN_mainform.Layer_name_ML_rectangle, 4, false);
                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                    poly_length = _AGEN_mainform.Poly3D.Length;
                    delete_centerlines();
                    zoom_to_object(_AGEN_mainform.Poly3D.ObjectId);
                    lista_del.Add(_AGEN_mainform.Poly3D.ObjectId);
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);


                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
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
                                    Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                    double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                    if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;
                                    double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                    _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                }
                            }
                        }

                        double dist1 = 0;
                        _AGEN_mainform.dt_sheet_index = Functions.Creaza_sheet_index_datatable_structure();
                        string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();

                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                        }
                        else
                        {
                            if (Scale1.Contains(":") == true)
                            {
                                Scale1 = Scale1.Replace("1:", "");
                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                                }
                            }
                            else
                            {
                                string inch = "\u0022";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                                }

                                inch = "\u0094";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                                }

                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                                }
                            }
                        }

                        double dist2 = dist1 + _AGEN_mainform.Match_distance;
                        bool Ultimul = false;
                        int Colorindex = 1;
                        bool Este_primul = true;
                        Point3d Last_pt = new Point3d();

                    l123:
                        Point3d Point1 = new Point3d();
                        Point3d Point2 = new Point3d();

                        if (radioButton_3D_station.Checked == true)
                        {
                            if (_AGEN_mainform.Poly3D.Length >= dist1)
                            {
                                Point1 = _AGEN_mainform.Poly3D.GetPointAtDist(dist1);
                            }
                            else
                            {
                                Point1 = _AGEN_mainform.Poly3D.StartPoint;
                                dist1 = 0;
                            }
                            if (_AGEN_mainform.Poly3D.Length >= dist2)
                            {
                                Point2 = _AGEN_mainform.Poly3D.GetPointAtDist(dist2);
                            }
                            else
                            {
                                Point2 = _AGEN_mainform.Poly3D.EndPoint;
                                Ultimul = true;
                                dist2 = _AGEN_mainform.Poly3D.Length;
                            }

                            Point1 = new Point3d(Point1.X, Point1.Y, 0);
                            Point2 = new Point3d(Point2.X, Point2.Y, 0);
                        }
                        else
                        {
                            if (_AGEN_mainform.Poly2D.Length >= dist1)
                            {
                                Point1 = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);
                            }
                            else
                            {
                                Point1 = _AGEN_mainform.Poly2D.StartPoint;
                                dist1 = 0;
                            }
                            if (_AGEN_mainform.Poly2D.Length >= dist2)
                            {
                                Point2 = _AGEN_mainform.Poly2D.GetPointAtDist(dist2);

                            }
                            else
                            {
                                Point2 = _AGEN_mainform.Poly2D.EndPoint;
                                Ultimul = true;
                                dist2 = _AGEN_mainform.Poly2D.Length;
                            }
                        }

                        Polyline Rectangle1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                        Rectangle1 = create_rectangle_Matchline(Point1, Point2, Colorindex);
                        Rectangle1.Layer = _AGEN_mainform.Layer_name_ML_rectangle;
                        Point3dCollection Col_int = new Point3dCollection();
                        Col_int = Functions.Intersect_on_both_operands(_AGEN_mainform.Poly2D, Rectangle1);

                        bool run_automatic = false;
                        if (checkBox_draw_automatic.Checked == true)
                        {
                            run_automatic = true;
                        }

                        if (run_automatic == true || Col_int.Count == 2)
                        {
                            BTrecord.AppendEntity(Rectangle1);
                            Trans1.AddNewlyCreatedDBObject(Rectangle1, true);
                            Trans1.TransactionManager.QueueForGraphicsFlush();
                            _AGEN_mainform.dt_sheet_index.Rows.Add();
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle1.ObjectId.Handle.Value.ToString();
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle1.GetPoint3dAt(0).X + Rectangle1.GetPoint3dAt(2).X) / 2;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle1.GetPoint3dAt(0).Y + Rectangle1.GetPoint3dAt(2).Y) / 2;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle1.GetPoint3dAt(1).X, Rectangle1.GetPoint3dAt(1).Y, Rectangle1.GetPoint3dAt(2).X, Rectangle1.GetPoint3dAt(2).Y) * 180 / Math.PI;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle1.GetPoint3dAt(1).DistanceTo(Rectangle1.GetPoint3dAt(2));
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle1.GetPoint3dAt(0).DistanceTo(Rectangle1.GetPoint3dAt(1));
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1.X;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1.Y;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2.X;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2.Y;

                            #region USA
                            if (_AGEN_mainform.COUNTRY == "USA")
                            {
                                _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = dist1;
                                _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = dist2;
                            }
                            #endregion

                            #region CANADA
                            if (_AGEN_mainform.COUNTRY == "CANADA")
                            {

                                double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(dist1);
                                double param2 = _AGEN_mainform.Poly3D.GetParameterAtDistance(dist2);

                                double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                double b1 = -1.23456;

                                double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                if (sta_csf1 != -1.23456)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = sta_csf1;
                                }

                                if (param2 > _AGEN_mainform.Poly2D.EndParam) param2 = _AGEN_mainform.Poly2D.EndParam;
                                double dist_2d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);

                                double b2 = -1.23456;
                                double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point2, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                if (sta_csf2 != -1.23456)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = sta_csf2;
                                }

                            }
                            #endregion

                            dist1 = dist2;
                            dist2 = dist2 + _AGEN_mainform.Match_distance;
                            Colorindex = Colorindex + 1;
                            if (Colorindex > 7) Colorindex = 1;

                            if (Ultimul == false)
                            {
                                if (radioButton_3D_station.Checked == false)
                                {
                                    if (Math.Round(_AGEN_mainform.Poly2D.Length, 0) <= Math.Round(dist2, 0))
                                    {
                                        dist2 = _AGEN_mainform.Poly2D.Length;
                                        Ultimul = true;
                                    }
                                }
                                else
                                {
                                    if (Math.Round(_AGEN_mainform.Poly3D.Length, 0) <= Math.Round(dist2, 0))
                                    {
                                        dist2 = _AGEN_mainform.Poly3D.Length - 0.0001;
                                        Ultimul = true;
                                    }
                                }

                                Este_primul = true;
                                goto l123;
                            }
                        }
                        else
                        {
                            Point3d Pointm1 = new Point3d();

                            if (dist1 > 0)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1m = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease pick start location:");


                                if (Este_primul == true)
                                {
                                    PP1m.AllowNone = false;
                                    Result_point_m1 = Editor1.GetPoint(PP1m);

                                    if (Result_point_m1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                    {
                                        Trans1.Commit();
                                        goto end1;
                                    }
                                    Pointm1 = Result_point_m1.Value;
                                    Last_pt = _AGEN_mainform.Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, false);
                                    if (radioButton_3D_station.Checked == true)
                                    {
                                        double paraml = _AGEN_mainform.Poly2D.GetParameterAtPoint(Last_pt);
                                        Last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(paraml);
                                    }

                                }
                            }

                            if (dist1 == 0)
                            {
                                Last_pt = _AGEN_mainform.Poly2D.GetPointAtParameter(0);
                                if (radioButton_3D_station.Checked == true)
                                {
                                    Last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(0);
                                }
                                Pointm1 = Last_pt;
                            }

                        labl1:

                            double dist1m;
                            double dist2m;

                            Point3d Point1m = new Point3d();
                            Point3d Point2m = new Point3d();

                            if (radioButton_3D_station.Checked == false)
                            {
                                Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2 Jig2 = new Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2 = Jig2.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly2D, Last_pt, 10);

                                if (Result_point_m2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                if (Este_primul == true)
                                {
                                    dist1m = _AGEN_mainform.Poly2D.GetDistAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, false));
                                }
                                else
                                {
                                    dist1m = _AGEN_mainform.Poly2D.GetDistAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, false));
                                }

                                Last_pt = _AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, false);
                                dist2m = _AGEN_mainform.Poly2D.GetDistAtPoint(Last_pt);

                                if (dist1m > dist2m)
                                {
                                    goto labl1;
                                }

                                Point1m = _AGEN_mainform.Poly2D.GetPointAtDist(dist1m);
                                Point2m = _AGEN_mainform.Poly2D.GetPointAtDist(dist2m);
                            }
                            else
                            {

                                Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2 Jig2 = new Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2 = Jig2.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly3D, _AGEN_mainform.Poly2D, Last_pt, 10);

                                if (Result_point_m2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                double Param1m;
                                if (Este_primul == true)
                                {
                                    Param1m = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Pointm1, Vector3d.ZAxis, false));
                                }
                                else
                                {
                                    Param1m = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Last_pt, Vector3d.ZAxis, false));
                                }

                                if (Param1m > _AGEN_mainform.Poly3D.EndParam) Param1m = _AGEN_mainform.Poly3D.EndParam;


                                dist1m = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param1m);

                                double Paramlast = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2.Value, Vector3d.ZAxis, false));
                                Last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(Paramlast);
                                dist2m = _AGEN_mainform.Poly3D.GetDistAtPoint(Last_pt);
                                if (dist1m > dist2m)
                                {
                                    goto labl1;
                                }
                                Point1m = _AGEN_mainform.Poly3D.GetPointAtDist(dist1m);
                                Point2m = _AGEN_mainform.Poly3D.GetPointAtDist(dist2m);
                            }

                            Polyline Rectangle2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                            Rectangle2 = create_rectangle_Matchline(Point1m, Point2m, Colorindex);
                            Rectangle2.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                            BTrecord.AppendEntity(Rectangle2);
                            Trans1.AddNewlyCreatedDBObject(Rectangle2, true);

                            Line Line1 = new Line(Rectangle2.GetPointAtParameter(2), Rectangle2.GetPointAtParameter(3));
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                            Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                            Jig1.AddEntity(Rectangle2);
                            Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                            if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Jig1.TransformEntities();
                            }

                            Trans1.TransactionManager.QueueForGraphicsFlush();

                            if (Este_primul == true)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                                {
                                    double M1_p = (double)_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1];
                                    double M2_p = (double)_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2];
                                    string ob_id = (string)_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle];
                                    ObjectId Oid = Functions.GetObjectId(ThisDrawing.Database, ob_id);
                                    Polyline PolyR = (Polyline)Trans1.GetObject(Oid, OpenMode.ForWrite);

                                    Point3d Point01 = new Point3d();
                                    Point3d Point02 = new Point3d();

                                    if (radioButton_3D_station.Checked == false)
                                    {
                                        Point01 = _AGEN_mainform.Poly2D.GetPointAtDist(M1_p);
                                        Point02 = _AGEN_mainform.Poly2D.GetPointAtDist(M2_p);
                                    }
                                    else
                                    {
                                        Point01 = _AGEN_mainform.Poly3D.GetPointAtDist(M1_p);
                                        Point02 = _AGEN_mainform.Poly3D.GetPointAtDist(M2_p);
                                    }

                                    Polyline Rectangle0 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                    Rectangle0 = create_rectangle_Matchline(Point01, Point1m, PolyR.ColorIndex);
                                    Rectangle0.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                                    BTrecord.AppendEntity(Rectangle0);
                                    Trans1.AddNewlyCreatedDBObject(Rectangle0, true);

                                    PolyR.Erase();
                                    Trans1.TransactionManager.QueueForGraphicsFlush();

                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle0.ObjectId.Handle.Value.ToString();
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle0.GetPoint3dAt(0).X + Rectangle0.GetPoint3dAt(2).X) / 2;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle0.GetPoint3dAt(0).Y + Rectangle0.GetPoint3dAt(2).Y) / 2;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle0.GetPoint3dAt(1).X, Rectangle0.GetPoint3dAt(1).Y, Rectangle0.GetPoint3dAt(2).X, Rectangle0.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle0.GetPoint3dAt(1).DistanceTo(Rectangle0.GetPoint3dAt(2));
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle0.GetPoint3dAt(0).DistanceTo(Rectangle0.GetPoint3dAt(1));
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point01.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point01.Y;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point1m.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point1m.Y;

                                    #region USA
                                    if (_AGEN_mainform.COUNTRY == "USA") _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = dist1m;
                                    #endregion

                                    #region CANADA
                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(dist1m);
                                        double dist_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                        double b1 = -1.23456;
                                        double sta_csf = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point02, dist_2d, _AGEN_mainform.dt_centerline, ref b1);
                                        if (sta_csf != -1.23456)
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = sta_csf;
                                        }


                                    }
                                    #endregion


                                }
                            }


                            _AGEN_mainform.dt_sheet_index.Rows.Add();
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle2.ObjectId.Handle.Value.ToString();
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle2.GetPoint3dAt(0).X + Rectangle2.GetPoint3dAt(2).X) / 2;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle2.GetPoint3dAt(0).Y + Rectangle2.GetPoint3dAt(2).Y) / 2;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle2.GetPoint3dAt(1).X, Rectangle2.GetPoint3dAt(1).Y, Rectangle2.GetPoint3dAt(2).X, Rectangle2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle2.GetPoint3dAt(1).DistanceTo(Rectangle2.GetPoint3dAt(2));
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle2.GetPoint3dAt(0).DistanceTo(Rectangle2.GetPoint3dAt(1));
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1m.X;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1m.Y;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2m.X;
                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2m.Y;

                            #region USA
                            if (_AGEN_mainform.COUNTRY == "USA")
                            {
                                _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = dist1m;
                                _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = dist2m;
                            }

                            #endregion

                            #region CANADA
                            if (_AGEN_mainform.COUNTRY == "CANADA")
                            {
                                double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(dist1m);


                                double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                double b1 = -1.23456;
                                double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1m, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                if (sta_csf1 != -1.23456)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = sta_csf1;
                                }

                                double param2 = _AGEN_mainform.Poly3D.GetParameterAtDistance(dist2m);
                                double dist_2d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);

                                double b2 = -1.23456;
                                double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point2m, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                if (sta_csf2 != -1.23456)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = sta_csf2;
                                }
                            }
                            #endregion



                            Colorindex = Colorindex + 1;
                            if (Colorindex > 7) Colorindex = 1;
                            Este_primul = false;
                            dist1 = dist2m;
                            dist2 = dist2m + _AGEN_mainform.Match_distance;

                            if (radioButton_3D_station.Checked == false)
                            {
                                if (Math.Round(dist1, 0) == Math.Round(_AGEN_mainform.Poly2D.Length, 0))
                                {
                                    goto l124;
                                }
                                if (Math.Round(dist2, 0) > Math.Round(_AGEN_mainform.Poly2D.Length, 0))
                                {
                                    dist2 = _AGEN_mainform.Poly2D.Length;
                                    Ultimul = true;
                                }
                            }
                            else
                            {
                                if (Math.Round(dist1, 0) == Math.Round(_AGEN_mainform.Poly3D.Length, 0))
                                {
                                    goto l124;
                                }
                                if (Math.Round(dist2, 0) > Math.Round(_AGEN_mainform.Poly3D.Length, 0))
                                {
                                    dist2 = _AGEN_mainform.Poly3D.Length - 0.0001;
                                    Ultimul = true;
                                }
                            }
                            goto l123;
                        }

                    l124: Editor1.WriteMessage("\nCommand:");
                        Trans1.Commit();
                    }
                }
            end1:
                if (_AGEN_mainform.dt_sheet_index != null)
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        Populate_data_table_matchline_file_names();

                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }
                            string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;
                            round_sheet_index_data_table(poly_length);
                            Append_ML_object_data();
                            dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                            dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                            dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                            dataGridView_sheet_index.EnableHeadersVisualStyles = false;
                            Functions.create_backup(fisier_si);
                            Populate_sheet_index_file(fisier_si);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                set_enable_true();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
            Ag.WindowState = FormWindowState.Normal;
        }

        private Polyline create_rectangle_Matchline(Point3d Point1, Point3d Point2, int cid)
        {
            Point1 = new Point3d(Point1.X, Point1.Y, 0);
            Point2 = new Point3d(Point2.X, Point2.Y, 0);
            Autodesk.AutoCAD.DatabaseServices.Line Line1R = new Autodesk.AutoCAD.DatabaseServices.Line(Point1, Point2);
            Point3d Point_distR = new Point3d();
            if (Line1R.Length > _AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale)
            {
                Point_distR = Line1R.GetPointAtDist(_AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale);
                Line1R.EndPoint = Point_distR;
            }
            else
            {
                Line1R.TransformBy(Matrix3d.Scaling((_AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale) / Line1R.Length, Line1R.StartPoint));
            }

            Line1R.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point1));
            Point3d Point_middler = new Point3d((Point1.X + Line1R.EndPoint.X) / 2, (Point1.Y + Line1R.EndPoint.Y) / 2, 0);

            Line1R.TransformBy(Matrix3d.Displacement(Point_middler.GetVectorTo(Point1)));
            Point3d Pt1r = new Point3d();
            Pt1r = Line1R.StartPoint;
            Point3d Pt2r = new Point3d();
            Pt2r = Line1R.EndPoint;
            Line1R.TransformBy(Matrix3d.Displacement(Point1.GetVectorTo(Point2)));

            Point3d Pt4r = new Point3d();
            Pt4r = Line1R.StartPoint;
            Point3d Pt3r = new Point3d();
            Pt3r = Line1R.EndPoint;

            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Pt1r.X, Pt1r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Pt2r.X, Pt2r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Pt3r.X, Pt3r.Y), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Pt4r.X, Pt4r.Y), 0, 0, 0);
            Poly1r.Closed = true;
            Poly1r.ColorIndex = cid;
            Poly1r.Elevation = 0;

            return Poly1r;
        }

        private void round_sheet_index_data_table(double poly_length)
        {
            if (_AGEN_mainform.dt_sheet_index != null)
            {
                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.dt_sheet_index.Rows[i][3] != DBNull.Value)
                        {
                            double St1 = Math.Round(Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][3]), _AGEN_mainform.round1);
                            double div1 = 10;
                            if (_AGEN_mainform.round1 == 1) div1 = 100;
                            if (_AGEN_mainform.round1 == 2) div1 = 1000;
                            if (_AGEN_mainform.round1 == 3) div1 = 10000;
                            if (St1 >= poly_length)
                            {
                                if (_AGEN_mainform.COUNTRY == "USA")
                                {
                                    St1 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                }
                            }

                            _AGEN_mainform.dt_sheet_index.Rows[i][3] = St1;

                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                            {
                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[i][5] = Math.Round(Functions.Station_equation_ofV2(St1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                }
                                else
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[i][5] = DBNull.Value;
                                }
                            }
                            else
                            {
                                _AGEN_mainform.dt_sheet_index.Rows[i][5] = DBNull.Value;
                            }
                        }
                        if (i < _AGEN_mainform.dt_sheet_index.Rows.Count - 1 && _AGEN_mainform.dt_sheet_index.Rows[i][4] != DBNull.Value)
                        {
                            double St2 = Math.Round(Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][4]), _AGEN_mainform.round1);

                            double div1 = 10;
                            if (_AGEN_mainform.round1 == 1) div1 = 100;
                            if (_AGEN_mainform.round1 == 2) div1 = 1000;
                            if (_AGEN_mainform.round1 == 3) div1 = 10000;

                            if (St2 >= poly_length)
                            {
                                if (_AGEN_mainform.COUNTRY == "USA")
                                {
                                    St2 = Math.Floor(Math.Round(poly_length * div1, _AGEN_mainform.round1 + 1)) / div1;
                                }

                            }
                            _AGEN_mainform.dt_sheet_index.Rows[i][4] = St2;
                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                            {
                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[i][6] = Math.Round(Functions.Station_equation_ofV2(St2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                }
                                else
                                {
                                    _AGEN_mainform.dt_sheet_index.Rows[i][6] = DBNull.Value;
                                }
                            }
                            else
                            {
                                _AGEN_mainform.dt_sheet_index.Rows[i][6] = DBNull.Value;
                            }
                        }
                    }
                }
            }
        }

        public void Populate_sheet_index_file(string File1)
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
                }

                else
                {
                    Workbook1 = Excel1.Workbooks.Open(File1);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {

                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_sheet_index, _AGEN_mainform.Start_row_Sheet_index, "General");
                    Functions.Create_header_sheet_index_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    if (System.IO.File.Exists(File1) == false)
                    {
                        Workbook1.SaveAs(File1);
                    }
                    else
                    {
                        Workbook1.Save();
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

        private void Populate_data_table_matchline_file_names()
        {
            if (_AGEN_mainform.dt_sheet_index != null)
            {
                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    string No_start = _AGEN_mainform.tpage_viewport_settings.get_start_number_from_text_box();
                    string Preffix = _AGEN_mainform.tpage_viewport_settings.get_prefix_name_from_text_box();
                    string Suffix = _AGEN_mainform.tpage_viewport_settings.get_suffix_name_from_text_box();

                    int Increment = 1;
                    if (Functions.IsNumeric(_AGEN_mainform.tpage_viewport_settings.get_increment_from_text_box()) == true)
                    {
                        Increment = Convert.ToInt32(_AGEN_mainform.tpage_viewport_settings.get_increment_from_text_box());
                    }


                    if (Functions.IsNumeric(No_start) == true)
                    {
                        int nr_start = Convert.ToInt32(No_start);
                        int old_nr = nr_start;

                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            string new_nr = old_nr.ToString();
                            if (i > 0) new_nr = (old_nr + Increment).ToString();
                            int len_no_start = No_start.Length;
                            int Len_new = new_nr.Length;
                            if (len_no_start > Len_new)
                            {
                                for (int j = Len_new; j < len_no_start; ++j)
                                {
                                    new_nr = "0" + new_nr;
                                }
                            }
                            string File_name = Preffix + new_nr + Suffix;
                            _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name] = File_name;
                            old_nr = Convert.ToInt32(new_nr);
                        }
                    }
                }
            }

        }


        private void Erase_viewports_templates()
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_name_VP_rectangle, "Agen_SheetIndex_VP");

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void Erase_northarrows()
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_odd, "Agen_Northarrow");
                        _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_even, "Agen_Northarrow");
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void Erase_matchline_blocks()
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_odd, "Agen_mlblocks");
                        _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_even, "Agen_mlblocks");
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        private void Create_ML_object_data()
        {

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

                            List1.Add("MMID");
                            List2.Add("ObjectID of the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DrawingNum");
                            List2.Add("Alignment_number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("BeginSta");
                            List2.Add("Matchline start");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("EndSta");
                            List2.Add("Matchline end");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Center_X");
                            List2.Add("X in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Center_Y");
                            List2.Add("Y in modelspace");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Rotation");
                            List2.Add("E-W viewport line rotation");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Width");
                            List2.Add("Matchline rectangle width");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Height");
                            List2.Add("Matchline rectangle height");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                            List1.Add("Type");
                            List2.Add("Type of drawing related to the rectangle");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Note1");
                            List2.Add("Notes");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("Version");
                            List2.Add("Version number");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("DateMod");
                            List2.Add("DateMod");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            List1.Add("SegmentName");
                            List2.Add("SegmentName");
                            List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Functions.Get_object_data_table("Agen_SheetIndex_ML", "Generated by AGEN", List1, List2, List3);

                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }


        public void delete_mtext_with_OD(string layer_name, string od_table_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        MText ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
                        if (ent1 != null)
                        {
                            if (ent1.Layer == layer_name)
                            {
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                bool delete1 = false;
                                if (Tables1.IsTableDefined(od_table_name) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[od_table_name];
                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1.Count > 0)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                            {
                                                if (delete1 == false)
                                                {
                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (Nume_field == "SegmentName")
                                                        {
                                                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                            if (segment1 == "not defined") segment1 = "";
                                                            if (Valoare1 == segment1)
                                                            {
                                                                delete1 = true;
                                                                i = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (delete1 == true)
                                    {
                                        ent1.UpgradeOpen();
                                        ent1.Erase();
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        public void delete_polyline_with_OD(string layer_name, string od_table_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        Polyline ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                        if (ent1 != null)
                        {
                            if (ent1.Layer == layer_name)
                            {
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                bool delete1 = false;
                                if (Tables1.IsTableDefined(od_table_name) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[od_table_name];
                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1.Count > 0)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                            {
                                                if (delete1 == false)
                                                {
                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (Nume_field == "SegmentName")
                                                        {
                                                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                            if (segment1 == "not defined") segment1 = "";
                                                            if (Valoare1 == segment1)
                                                            {
                                                                delete1 = true;
                                                                i = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (delete1 == true)
                                    {
                                        ent1.UpgradeOpen();
                                        ent1.Erase();
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        private void Append_ML_object_data()
        {

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

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            delete_mtext_with_OD(_AGEN_mainform.Layer_name_ML_rectangle, "Agen_SheetIndex_ML");




                            for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                            {

                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                string ObjID = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_handle].ToString();

                                ObjectId id_poly = Functions.GetObjectId(ThisDrawing.Database, ObjID);
                                if (id_poly != ObjectId.Null)
                                {
                                    Lista_val.Add(ObjID);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Polyline Poly1 = Trans1.GetObject(id_poly, OpenMode.ForWrite) as Polyline;

                                    string nume_dwg = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();

                                    MText Mt1_label = new MText();
                                    Mt1_label.Contents = nume_dwg;

                                    Mt1_label.TextHeight = _AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale / 10;

                                    if (_AGEN_mainform.Left_to_Right == true)
                                    {
                                        Mt1_label.Rotation = Functions.GET_Bearing_rad(Poly1.GetPointAtParameter(1).X, Poly1.GetPointAtParameter(1).Y, Poly1.GetPointAtParameter(2).X, Poly1.GetPointAtParameter(2).Y);
                                        Mt1_label.Attachment = AttachmentPoint.BottomLeft;
                                        Mt1_label.Location = Poly1.GetPointAtParameter(1);
                                    }
                                    else
                                    {
                                        Mt1_label.Rotation = Functions.GET_Bearing_rad(Poly1.GetPointAtParameter(2).X, Poly1.GetPointAtParameter(2).Y, Poly1.GetPointAtParameter(1).X, Poly1.GetPointAtParameter(1).Y);
                                        Mt1_label.Attachment = AttachmentPoint.BottomLeft;
                                        Mt1_label.Location = Poly1.GetPointAtParameter(3);
                                    }

                                    Mt1_label.ColorIndex = Poly1.ColorIndex;
                                    Mt1_label.Layer = _AGEN_mainform.Layer_name_ML_rectangle;
                                    BTrecord.AppendEntity(Mt1_label);
                                    Trans1.AddNewlyCreatedDBObject(Mt1_label, true);



                                    Lista_val.Add(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString());
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Width]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add((double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Height]);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add("Alignment Sheet");
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add("");
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add("");
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                    if (segment1 == "not defined") segment1 = "";

                                    Lista_val.Add(segment1);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                                    Functions.Populate_object_data_table_from_handle_string(Tables1, ObjID, "Agen_SheetIndex_ML", Lista_val, Lista_type);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, Mt1_label.ObjectId, "Agen_SheetIndex_ML", Lista_val, Lista_type);

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

        }

        private void button_draw_Viewport_templates_Click(object sender, EventArgs e)
        {

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)

                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }


            string Fisier_si = "";
            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }




                Fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                if (System.IO.File.Exists(Fisier_si) == false)
                {
                    set_enable_true();
                    MessageBox.Show("No sheet index file loaded");
                    return;
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            set_enable_false();
            _AGEN_mainform.tpage_processing.Show();
            Erase_viewports_templates();

            Create_VP_object_data();


            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_AGEN_mainform.Layer_name_VP_rectangle, 7, false);

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {



                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        if (_AGEN_mainform.dt_sheet_index != null)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {

                                _AGEN_mainform.Data_table_Main_VP = new System.Data.DataTable();
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_handle, typeof(string));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_x, typeof(double));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_y, typeof(double));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_rot, typeof(double));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_Width, typeof(double));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_Height, typeof(double));
                                _AGEN_mainform.Data_table_Main_VP.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));
                                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {
                                    _AGEN_mainform.Data_table_Main_VP.Rows.Add();

                                }
                                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {
                                    double X = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x];
                                    double Y = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y];
                                    double Rotation1 = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot];

                                    string Objectid_string = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_handle].ToString();

                                    ObjectId ObjectID1 = Functions.GetObjectId(ThisDrawing.Database, Objectid_string);

                                    int CI = 256;
                                    if (ObjectID1 != null)
                                    {
                                        Entity Ent1 = (Entity)Trans1.GetObject(ObjectID1, OpenMode.ForRead);
                                        if (Ent1 != null)
                                        {
                                            CI = Ent1.ColorIndex;
                                        }
                                    }
                                    Polyline Poly1 = creaza_rectangle_from_one_point(new Point3d(X, Y, 0), Rotation1 * Math.PI / 180, _AGEN_mainform.Vw_width / _AGEN_mainform.Vw_scale, _AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale, CI);
                                    Poly1.Layer = _AGEN_mainform.Layer_name_VP_rectangle;
                                    BTrecord.AppendEntity(Poly1);
                                    Trans1.AddNewlyCreatedDBObject(Poly1, true);

                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_handle] = Poly1.ObjectId.Handle.Value.ToString();
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_x] = X;
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_y] = Y;
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_rot] = Rotation1;
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_Width] = _AGEN_mainform.Vw_width;
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_Height] = _AGEN_mainform.Vw_height;
                                    _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();

                                }

                                label_vp.Visible = true;
                            }
                        }

                        Editor1.WriteMessage("\nCommand:");

                        Trans1.Commit();
                    }
                }




                if (_AGEN_mainform.Data_table_Main_VP != null)
                {
                    if (_AGEN_mainform.Data_table_Main_VP.Rows.Count > 0)
                    {
                        Append_VP_object_data();


                    }
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();

        }

        public Polyline creaza_rectangle_from_one_point(Point3d Point1, double Rotation_rad, double Width1, double Height1, int cid)
        {
            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Point1.X - Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Point1.X - Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Point1.X + Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Point1.X + Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);


            Poly1r.Closed = true;
            Poly1r.ColorIndex = cid;

            Poly1r.TransformBy(Matrix3d.Rotation(Rotation_rad, Vector3d.ZAxis, Point1));

            return Poly1r;
        }

        private void Create_VP_object_data()
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

                        List1.Add("MMID");
                        List2.Add("ObjectID of the rectangle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("DrawingNum");
                        List2.Add("Alignment_number");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Center_X");
                        List2.Add("X in modelspace");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Center_Y");
                        List2.Add("Y in modelspace");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Rotation");
                        List2.Add("E-W viewport line rotation");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Width");
                        List2.Add("Matchline rectangle width");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Height");
                        List2.Add("Matchline rectangle height");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                        List1.Add("Type");
                        List2.Add("Type of drawing related to the rectangle");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Note1");
                        List2.Add("Notes");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("Version");
                        List2.Add("Version number");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("DateMod");
                        List2.Add("DateMod");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add("SegmentName");
                        List2.Add("SegmentName");
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("Agen_SheetIndex_VP", "Generated by AGEN", List1, List2, List3);

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Append_VP_object_data()
        {
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
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            for (int i = 0; i < _AGEN_mainform.Data_table_Main_VP.Rows.Count; ++i)
                            {

                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                string ObjID = _AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_handle].ToString();

                                Lista_val.Add(ObjID);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Entity Ent1 = (Entity)Trans1.GetObject(Functions.GetObjectId(ThisDrawing.Database, ObjID), OpenMode.ForWrite);

                                Lista_val.Add(_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_dwg_name].ToString());
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add((double)_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_x]);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                Lista_val.Add((double)_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_y]);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                Lista_val.Add((double)_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_rot]);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                Lista_val.Add((double)_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_Width]);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                Lista_val.Add((double)_AGEN_mainform.Data_table_Main_VP.Rows[i][_AGEN_mainform.Col_Height]);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                Lista_val.Add("Alignment Sheet");
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add("");
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add("");
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                if (segment1 == "not defined") segment1 = "";
                                Lista_val.Add(segment1);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                                Functions.Populate_object_data_table_from_handle_string(Tables1, ObjID, "Agen_SheetIndex_VP", Lista_val, Lista_type);
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

        }




        private void button_Fill_gaps_Click(object sender, EventArgs e)
        {
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


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)

                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }

            if (_AGEN_mainform.dt_centerline == null)
            {
                MessageBox.Show("you do not have picked the centerline\r\noperation aborted");

                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();


                return;
            }

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("the centerline file d\r\noperation aborted");


                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();


                return;
            }



            if (_AGEN_mainform.dt_sheet_index == null)
            {
                MessageBox.Show("the sheet index table is null\r\noperation aborted");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("No data for sheet indexes found\r\noperation aborted");
                return;
            }


            Ag.WindowState = FormWindowState.Minimized;

            set_enable_false();

            Erase_viewports_templates();

            if (Functions.IsNumeric(TextBox_matchline_length.Text) == true)
            {
                _AGEN_mainform.Match_distance = Convert.ToDouble(TextBox_matchline_length.Text);
            }

            string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();
            if (Functions.IsNumeric(Scale1) == true)
            {
                _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
            }
            else
            {
                if (Scale1.Contains(":") == true)
                {
                    Scale1 = Scale1.Replace("1:", "");
                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                    }
                }
                else
                {
                    string inch = "\u0022";

                    if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                    {
                        Scale1 = Scale1.Replace("1" + inch + "=", "");
                        Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                    }

                    inch = "\u0094";

                    if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                    {
                        Scale1 = Scale1.Replace("1" + inch + "=", "");
                        Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                    }

                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                    }
                }
            }

            double poly_length = 0;
            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_AGEN_mainform.Layer_name_ML_rectangle, 4, false);

                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                    delete_centerlines();
                    lista_del.Add(_AGEN_mainform.Poly3D.ObjectId);

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        string Path_toCL = "";
                        string Fisier_si = "";
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            Path_toCL = ProjF + _AGEN_mainform.cl_excel_name;

                            if (System.IO.File.Exists(Path_toCL) == false)
                            {
                                set_enable_true();
                                MessageBox.Show("No centerline file loaded");
                                return;
                            }

                            Fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                            if (System.IO.File.Exists(Fisier_si) == false)
                            {
                                set_enable_true();
                                MessageBox.Show("No sheet index file loaded");
                                return;
                            }

                        }
                        else
                        {
                            set_enable_true();
                            MessageBox.Show("the project folder does not exist");
                            return;
                        }






                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        //Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                        #region USA
                        if (_AGEN_mainform.COUNTRY == "USA")
                        {


                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
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


                                        Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                        double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                        if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;
                                        double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                    }
                                }
                            }
                        }

                        #endregion



                        List<int> List1 = new List<int>();

                        bool deleted = false;

                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {


                            System.Data.DataTable dt_si = Functions.Creaza_sheet_index_datatable_structure();
                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline rectangle_ml = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                if (rectangle_ml != null)
                                {
                                    if (rectangle_ml.Layer == _AGEN_mainform.Layer_name_ML_rectangle)
                                    {
                                        if (rectangle_ml.NumberOfVertices == 4)
                                        {
                                            _AGEN_mainform.Poly2D.Elevation = rectangle_ml.Elevation;
                                            Point3dCollection col_int = Functions.Intersect_on_both_operands(rectangle_ml, _AGEN_mainform.Poly2D);

                                            Polyline p1 = new Polyline();
                                            p1.AddVertexAt(0, rectangle_ml.GetPoint2dAt(0), 0, 0, 0);
                                            p1.AddVertexAt(1, rectangle_ml.GetPoint2dAt(1), 0, 0, 0);
                                            p1.Elevation = rectangle_ml.Elevation;

                                            Polyline p2 = new Polyline();
                                            p2.AddVertexAt(0, rectangle_ml.GetPoint2dAt(2), 0, 0, 0);
                                            p2.AddVertexAt(1, rectangle_ml.GetPoint2dAt(3), 0, 0, 0);
                                            p2.Elevation = rectangle_ml.Elevation;

                                            Point3dCollection col_int1 = Functions.Intersect_on_both_operands(p1, _AGEN_mainform.Poly2D);
                                            Point3dCollection col_int2 = Functions.Intersect_on_both_operands(p2, _AGEN_mainform.Poly2D);

                                            bool exista_rectangle = false;
                                            if (col_int.Count >= 2)
                                            {
                                                exista_rectangle = true;
                                            }

                                            if (col_int.Count == 1)
                                            {
                                                Point3d point_on_col = col_int[0];
                                                if (col_int1.Count == 0)
                                                {
                                                    Point3d point_onP1 = p1.GetClosestPointTo(_AGEN_mainform.Poly2D.StartPoint, Vector3d.ZAxis, false);
                                                    double dist_at_P1 = Math.Pow(Math.Pow(_AGEN_mainform.Poly2D.StartPoint.X - point_onP1.X, 2) + Math.Pow(_AGEN_mainform.Poly2D.StartPoint.Y - point_onP1.Y, 2), 0.5);
                                                    if (dist_at_P1 < 0.1)
                                                    {
                                                        exista_rectangle = true;
                                                        col_int.Add(_AGEN_mainform.Poly2D.StartPoint);
                                                    }
                                                }
                                                else
                                                {
                                                    Point3d point_on_col1 = col_int1[0];
                                                    double dist_at_P1 = Math.Pow(Math.Pow(point_on_col.X - point_on_col1.X, 2) + Math.Pow(point_on_col.Y - point_on_col1.Y, 2), 0.5);
                                                    if (dist_at_P1 > 1)
                                                    {
                                                        exista_rectangle = true;
                                                        col_int.Add(point_on_col1);
                                                    }
                                                }


                                                if (col_int2.Count == 0)
                                                {
                                                    Point3d point_onP2 = p2.GetClosestPointTo(_AGEN_mainform.Poly2D.StartPoint, Vector3d.ZAxis, false);
                                                    double dist_at_P2 = Math.Pow(Math.Pow(_AGEN_mainform.Poly2D.StartPoint.X - point_onP2.X, 2) + Math.Pow(_AGEN_mainform.Poly2D.StartPoint.Y - point_onP2.Y, 2), 0.5);
                                                    if (dist_at_P2 < 0.1)
                                                    {
                                                        exista_rectangle = true;
                                                        col_int.Add(_AGEN_mainform.Poly2D.StartPoint);
                                                    }
                                                }
                                                else
                                                {
                                                    Point3d point_on_col2 = col_int2[0];
                                                    double dist_at_P2 = Math.Pow(Math.Pow(point_on_col.X - point_on_col2.X, 2) + Math.Pow(point_on_col.Y - point_on_col2.Y, 2), 0.5);
                                                    if (dist_at_P2 > 1)
                                                    {
                                                        exista_rectangle = true;
                                                        col_int.Add(point_on_col2);
                                                    }
                                                }
                                            }


                                            if (exista_rectangle == true)
                                            {

                                                if (col_int1.Count > 0 & col_int2.Count > 0)
                                                {

                                                    for (int m = 0; m < col_int1.Count; ++m)
                                                    {
                                                        double prm1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(col_int1[m], Vector3d.ZAxis, false));

                                                        if (prm1 >= _AGEN_mainform.Poly2D.EndParam)
                                                        {
                                                            prm1 = _AGEN_mainform.Poly3D.EndParam;
                                                        }
                                                        if (prm1 > _AGEN_mainform.Poly3D.EndParam) prm1 = _AGEN_mainform.Poly3D.EndParam;

                                                        double d1 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(prm1);
                                                        dt_si.Rows.Add();
                                                        dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_handle] = rectangle_ml.ObjectId.Handle.Value.ToString();



                                                        #region USA
                                                        if (_AGEN_mainform.COUNTRY == "USA") dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_M1] = Math.Round(d1, _AGEN_mainform.round1);
                                                        #endregion

                                                        #region CANADA
                                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                                        {
                                                            double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(d1);
                                                            double dist_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                                            double b1 = -1.23456;

                                                            double sta_csf = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, col_int1[m], dist_2d, _AGEN_mainform.dt_centerline, ref b1);
                                                            if (sta_csf != -1.23456) dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_M1] = Math.Round(sta_csf, _AGEN_mainform.round1);
                                                        }
                                                        #endregion



                                                        dt_si.Rows[dt_si.Rows.Count - 1]["X_Beg"] = col_int1[m].X;
                                                        dt_si.Rows[dt_si.Rows.Count - 1]["Y_Beg"] = col_int1[m].Y;
                                                    }


                                                    for (int m = 0; m < col_int2.Count; ++m)
                                                    {
                                                        double prm2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(col_int2[m], Vector3d.ZAxis, false));

                                                        if (prm2 >= _AGEN_mainform.Poly2D.EndParam)
                                                        {
                                                            prm2 = _AGEN_mainform.Poly3D.EndParam;
                                                        }

                                                        if (prm2 > _AGEN_mainform.Poly3D.EndParam) prm2 = _AGEN_mainform.Poly3D.EndParam;
                                                        double d2 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(prm2);
                                                        dt_si.Rows.Add();
                                                        dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_handle] = rectangle_ml.ObjectId.Handle.Value.ToString();
                                                        #region USA
                                                        if (_AGEN_mainform.COUNTRY == "USA") dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_M1] = Math.Round(d2, _AGEN_mainform.round1);
                                                        #endregion

                                                        #region CANADA
                                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                                        {
                                                            double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(d2);
                                                            double dist_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                                            double b1 = -1.23456;
                                                            double sta_csf = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, col_int2[m], dist_2d, _AGEN_mainform.dt_centerline, ref b1);
                                                            if (sta_csf != -1.23456) dt_si.Rows[dt_si.Rows.Count - 1][_AGEN_mainform.Col_M1] = Math.Round(sta_csf, _AGEN_mainform.round1);
                                                        }
                                                        #endregion

                                                        dt_si.Rows[dt_si.Rows.Count - 1]["X_End"] = col_int2[m].X;
                                                        dt_si.Rows[dt_si.Rows.Count - 1]["Y_End"] = col_int2[m].Y;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                            }

                            dt_si = Functions.Sort_data_table(dt_si, _AGEN_mainform.Col_M1);

                            if (_AGEN_mainform.dt_sheet_index.Columns.Contains("X_Beg") == false ||
                                _AGEN_mainform.dt_sheet_index.Columns.Contains("Y_Beg") == false ||
                                _AGEN_mainform.dt_sheet_index.Columns.Contains("X_End") == false ||
                                _AGEN_mainform.dt_sheet_index.Columns.Contains("Y_End") == false)
                            {
                                _AGEN_mainform.tpage_processing.Hide();

                                Ag.WindowState = FormWindowState.Normal;

                                MessageBox.Show("first you have to use recover button");


                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                            }

                            for (int i = _AGEN_mainform.dt_sheet_index.Rows.Count - 1; i >= 0; --i)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[i]["X_Beg"] != DBNull.Value &&
                                    _AGEN_mainform.dt_sheet_index.Rows[i]["Y_Beg"] != DBNull.Value &&
                                    _AGEN_mainform.dt_sheet_index.Rows[i]["X_End"] != DBNull.Value &&
                                    _AGEN_mainform.dt_sheet_index.Rows[i]["Y_End"] != DBNull.Value)
                                {
                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["X_Beg"]);
                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["Y_Beg"]);
                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["X_End"]);
                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["Y_End"]);

                                    bool pt1_found = false;
                                    bool pt2_found = false;

                                    for (int j = 0; j < dt_si.Rows.Count; ++j)
                                    {


                                        double xp1 = -1.23;
                                        double yp1 = -1.23;

                                        if (dt_si.Rows[j]["X_Beg"] != DBNull.Value & dt_si.Rows[j]["Y_Beg"] != DBNull.Value && pt1_found == false)
                                        {
                                            xp1 = Convert.ToDouble(dt_si.Rows[j]["X_Beg"]);
                                            yp1 = Convert.ToDouble(dt_si.Rows[j]["Y_Beg"]);
                                        }

                                        else if (dt_si.Rows[j]["X_End"] != DBNull.Value & dt_si.Rows[j]["Y_End"] != DBNull.Value && pt2_found == false)
                                        {
                                            xp1 = Convert.ToDouble(dt_si.Rows[j]["X_End"]);
                                            yp1 = Convert.ToDouble(dt_si.Rows[j]["Y_End"]);
                                        }
                                        double d1 = Math.Pow(Math.Pow(x1-xp1, 2)+ Math.Pow(y1-yp1, 2), 0.5);
                                        double M1 = -1;
                                        if (dt_si.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value)
                                        {
                                           M1 = Convert.ToDouble(dt_si.Rows[j][_AGEN_mainform.Col_M1]);
                                        }

                                        if (d1 <= 0.01 && pt1_found == false)
                                        {
                                            pt1_found = true;
                                        }

                                        if (new Point2d(x2, y2).GetDistanceTo(new Point2d(xp1, yp1)) <= 0.01 && pt2_found == false)
                                        {
                                            pt2_found = true;
                                        }
                                    }

                                    if (pt1_found == false || pt2_found == false)
                                    {
                                        _AGEN_mainform.dt_sheet_index.Rows[i].Delete();
                                        deleted = true;
                                    }
                                    else
                                    {
                                        Point3d pt_on_line1 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x1, y1, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                        Point3d pt_on_line2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);

                                        double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_line1);
                                        double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_line2);

                                        if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;
                                        if (param2 > _AGEN_mainform.Poly3D.EndParam) param2 = _AGEN_mainform.Poly3D.EndParam;

                                        double d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                        double d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);

                                        double m1 = Math.Round(_AGEN_mainform.Poly3D.GetDistanceAtParameter(param1), _AGEN_mainform.round1);
                                        double m2 = Math.Round(_AGEN_mainform.Poly3D.GetDistanceAtParameter(param2), _AGEN_mainform.round1);

                                        #region USA
                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1] = m1;
                                            _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2] = m2;
                                        }
                                        #endregion

                                        #region CANADA
                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                        {
                                            double b1 = -1.23456;
                                            double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, pt_on_line1, d1, _AGEN_mainform.dt_centerline, ref b1);
                                            _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1] = Math.Round(sta_csf1, _AGEN_mainform.round1);

                                            double b2 = -1.23456;
                                            double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, pt_on_line2, d2, _AGEN_mainform.dt_centerline, ref b2);
                                            _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2] = Math.Round(sta_csf2, _AGEN_mainform.round1);
                                        }
                                        #endregion
                                    }
                                }
                                else
                                {
                                    _AGEN_mainform.tpage_processing.Hide();
                                    Ag.WindowState = FormWindowState.Normal;
                                    MessageBox.Show("first you have to use recover button");
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                }


                            }
                        }


                        double first_matchline_value = 0;
                        if (_AGEN_mainform.COUNTRY == "CANADA")
                        {
                            if (_AGEN_mainform.dt_centerline != null && _AGEN_mainform.dt_centerline.Rows.Count > 0)
                            {
                                if (_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                {
                                    first_matchline_value = Math.Round(Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_3DSta]), _AGEN_mainform.round1);
                                }
                            }
                        }

                        List<int> List2 = new List<int>();



                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {
                            for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                            {
                                double m1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1]);
                                double m2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2]);
                                if (m1 != first_matchline_value)
                                {
                                    List2.Add(i);
                                    deleted = true;
                                }
                                first_matchline_value = m2;
                            }
                        }
                        else
                        {
                            deleted = true;
                        }


                        if (radioButton_3D_station.Checked == false)
                        {
                            poly_length = _AGEN_mainform.Poly2D.Length;
                        }
                        else
                        {
                            poly_length = _AGEN_mainform.Poly3D.Length;

                        }

                        Point3d last_pt = new Point3d(1.234, 1.234, 1.234);
                        double dist1 = 0;
                        #region USA
                        if (_AGEN_mainform.COUNTRY == "USA")
                        {
                            double lastM = 0;
                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {
                                lastM = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2]);
                            }
                            if (Math.Abs(lastM - poly_length) > 1)
                            {
                                deleted = true;
                                last_pt = _AGEN_mainform.Poly2D.GetPointAtDist(lastM);
                                dist1 = lastM;
                            }
                        }
                        #endregion

                        #region CANADA
                        if (_AGEN_mainform.COUNTRY == "CANADA")
                        {
                            double x2 = 0;
                            double y2 = 0;
                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {
                                x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"]);
                                y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"]);
                            }
                            else
                            {
                                x2 = _AGEN_mainform.Poly2D.StartPoint.X;
                                y2 = _AGEN_mainform.Poly2D.StartPoint.Y;
                            }


                            if (Math.Abs(_AGEN_mainform.Poly2D.GetDistAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false)) - _AGEN_mainform.Poly2D.Length) > 1)
                            {
                                deleted = true;
                                last_pt = new Point3d(x2, y2, 0);
                            }
                        }
                        #endregion



                        if (deleted == false)
                        {
                            MessageBox.Show("All the matcline rectangles from sheet index are present.\r\n" +
                                            "This function works when some matchline rectangles are missing\r\n" +
                                            "Operation aborted");
                            set_enable_true();
                            Ag.WindowState = FormWindowState.Normal;
                            return;
                        }

                        int Colorindex = 1;

                        //Point3d last_pt = new Point3d();

                        double dist2 = 0;
                        double Next_matchline = 0;


                        if (List2.Count > 0)
                        {
                            #region lista2.count>0
                            for (int i = 0; i < List2.Count; ++i)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_Beg"] != DBNull.Value &&
                                    _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_Beg"] != DBNull.Value)
                                {
                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_Beg"]);
                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_Beg"]);
                                    Point3d pt_on_line1 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x1, y1, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                    double Param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_line1);
                                    if (Param1 > _AGEN_mainform.Poly3D.EndParam) Param1 = _AGEN_mainform.Poly3D.EndParam;
                                    Next_matchline = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param1);

                                    if (List2[i] != 0)
                                    {
                                        #region USA
                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            dist1 = (double)_AGEN_mainform.dt_sheet_index.Rows[List2[i] - 1][_AGEN_mainform.Col_M2];
                                        }
                                        #endregion

                                        #region CANADA
                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                        {
                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i] - 1]["X_End"]);
                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i] - 1]["Y_End"]);
                                            Point3d pt_on_line2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                            double Param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_line2);
                                            if (Param2 > _AGEN_mainform.Poly3D.EndParam) Param2 = _AGEN_mainform.Poly3D.EndParam;
                                            dist1 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param2);
                                        }
                                        #endregion
                                    }

                                l1234:
                                    Point3d Point1m = new Point3d();
                                    Point3d Point2m = new Point3d();

                                    if (radioButton_3D_station.Checked == false)
                                    {
                                        #region 2D
                                        double div1 = 10;
                                        if (_AGEN_mainform.round1 == 1) div1 = 100;
                                        if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                        if (_AGEN_mainform.round1 == 3) div1 = 10000;



                                        if (dist1 >= _AGEN_mainform.Poly2D.Length)
                                        {
                                            dist1 = Math.Floor(Math.Round(_AGEN_mainform.Poly2D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                        }


                                        last_pt = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);


                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                        zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                        Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2();
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly2D, last_pt, 10);

                                        if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                        {
                                            Trans1.Commit();
                                            goto end1;
                                        }

                                        last_pt = _AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false);

                                        if (_AGEN_mainform.Poly2D.EndPoint.DistanceTo(last_pt) < 0.0009)
                                        {
                                            last_pt = _AGEN_mainform.Poly2D.EndPoint;
                                        }

                                        dist2 = _AGEN_mainform.Poly2D.GetDistAtPoint(last_pt);

                                        if (Math.Round(dist1, 0) > Math.Round(dist2, 0))
                                        {
                                            goto l1234;
                                        }


                                        Point1m = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);

                                        Point2m = _AGEN_mainform.Poly2D.GetPointAtDist(dist2);
                                        #endregion


                                    }
                                    else
                                    {
                                        double div1 = 10;
                                        if (_AGEN_mainform.round1 == 1) div1 = 100;
                                        if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                        if (_AGEN_mainform.round1 == 3) div1 = 10000;


                                        if (dist1 >= _AGEN_mainform.Poly3D.Length)
                                        {
                                            dist1 = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                        }

                                        last_pt = _AGEN_mainform.Poly3D.GetPointAtDist(dist1);

                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                        zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();


                                        Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2();

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly3D, _AGEN_mainform.Poly2D, last_pt, 10);

                                        if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                        {
                                            Trans1.Commit();
                                            goto end1;
                                        }

                                        double p_l = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false));
                                        if (Math.Round(p_l, 4) == Math.Round(_AGEN_mainform.Poly2D.EndParam, 4))
                                        {
                                            p_l = _AGEN_mainform.Poly2D.EndParam;
                                        }


                                        last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(p_l);

                                        dist2 = _AGEN_mainform.Poly3D.GetDistAtPoint(last_pt);

                                        if (Math.Round(dist1, 0) > Math.Round(dist2, 0))
                                        {
                                            goto l1234;
                                        }


                                        Point1m = _AGEN_mainform.Poly3D.GetPointAtDist(dist1);

                                        Point2m = _AGEN_mainform.Poly3D.GetPointAtDist(dist2);
                                    }



                                    Polyline Rectangle_ML1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                    Rectangle_ML1 = create_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                    Rectangle_ML1.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                                    BTrecord.AppendEntity(Rectangle_ML1);
                                    Trans1.AddNewlyCreatedDBObject(Rectangle_ML1, true);


                                    Line Line1 = new Line(Rectangle_ML1.GetPointAtParameter(2), Rectangle_ML1.GetPointAtParameter(3));
                                    Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                    Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                    Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                    Jig1.AddEntity(Rectangle_ML1);
                                    Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                    if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Jig1.TransformEntities();
                                    }

                                    Trans1.TransactionManager.QueueForGraphicsFlush();


                                    System.Data.DataRow Row1 = _AGEN_mainform.dt_sheet_index.NewRow();


                                    Row1[_AGEN_mainform.Col_handle] = Rectangle_ML1.ObjectId.Handle.Value.ToString();
                                    Row1[_AGEN_mainform.Col_x] = (Rectangle_ML1.GetPoint3dAt(0).X + Rectangle_ML1.GetPoint3dAt(2).X) / 2;
                                    Row1[_AGEN_mainform.Col_y] = (Rectangle_ML1.GetPoint3dAt(0).Y + Rectangle_ML1.GetPoint3dAt(2).Y) / 2;
                                    Row1[_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle_ML1.GetPoint3dAt(1).X, Rectangle_ML1.GetPoint3dAt(1).Y, Rectangle_ML1.GetPoint3dAt(2).X, Rectangle_ML1.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                    Row1[_AGEN_mainform.Col_Width] = Rectangle_ML1.GetPoint3dAt(1).DistanceTo(Rectangle_ML1.GetPoint3dAt(2));
                                    Row1[_AGEN_mainform.Col_Height] = Rectangle_ML1.GetPoint3dAt(0).DistanceTo(Rectangle_ML1.GetPoint3dAt(1));
                                    Row1["X_Beg"] = Point1m.X;
                                    Row1["Y_Beg"] = Point1m.Y;
                                    Row1["X_End"] = Point2m.X;
                                    Row1["Y_End"] = Point2m.Y;

                                    #region USA
                                    if (_AGEN_mainform.COUNTRY == "USA")
                                    {
                                        Row1[_AGEN_mainform.Col_M1] = dist1;
                                        Row1[_AGEN_mainform.Col_M2] = dist2;
                                    }
                                    #endregion

                                    #region CANADA
                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        double param1 = _AGEN_mainform.Poly3D.GetParameterAtPoint(Point1m);
                                        double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                        double b1 = -1.23456;
                                        double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1m, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                        if (sta_csf1 != -1.23456) Row1[_AGEN_mainform.Col_M1] = Math.Round(sta_csf1, _AGEN_mainform.round1);

                                        double param2 = _AGEN_mainform.Poly3D.GetParameterAtPoint(Point2m);
                                        double dist_2d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);
                                        double b2 = -1.23456;
                                        double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point2m, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                        if (sta_csf2 != -1.23456) Row1[_AGEN_mainform.Col_M2] = Math.Round(sta_csf2, _AGEN_mainform.round1);
                                    }
                                    #endregion

                                    _AGEN_mainform.dt_sheet_index.Rows.InsertAt(Row1, List2[i]);

                                    for (int j = i; j < List2.Count; ++j)
                                    {
                                        List2[j] = List2[j] + 1;
                                    }


                                    Colorindex = Colorindex + 1;
                                    if (Colorindex > 7) Colorindex = 1;


                                    dist1 = dist2;

                                    if (Math.Round(dist2, 0) < Math.Round(Next_matchline, 0))
                                    {
                                        goto l1234;
                                    }
                                    else
                                    {


                                        Point3d Point1mn = new Point3d();
                                        Point3d Point2mn = new Point3d();

                                        if (_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_End"] != DBNull.Value &&
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_End"] != DBNull.Value)
                                        {
                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_End"]);
                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_End"]);

                                            Point3d Point_online2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                            double param22 = _AGEN_mainform.Poly2D.GetParameterAtPoint(Point_online2);
                                            if (param22 > _AGEN_mainform.Poly3D.EndParam) param22 = _AGEN_mainform.Poly3D.EndParam;
                                            double dist2mn = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param22);



                                            double div1 = 10;
                                            if (_AGEN_mainform.round1 == 1) div1 = 100;
                                            if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                            if (_AGEN_mainform.round1 == 3) div1 = 10000;




                                            if (radioButton_3D_station.Checked == false)
                                            {
                                                #region 2D
                                                if (dist1 >= _AGEN_mainform.Poly2D.Length)
                                                {
                                                    dist1 = Math.Floor(Math.Round(_AGEN_mainform.Poly2D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                                }

                                                if (dist2mn >= _AGEN_mainform.Poly2D.Length)
                                                {
                                                    dist2mn = Math.Floor(Math.Round(_AGEN_mainform.Poly2D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                                }

                                                Point1mn = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);
                                                Point2mn = _AGEN_mainform.Poly2D.GetPointAtDist(dist2mn);
                                                #endregion

                                            }
                                            else
                                            {

                                                if (dist1 >= _AGEN_mainform.Poly3D.Length)
                                                {
                                                    dist1 = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                                }

                                                if (dist2mn >= _AGEN_mainform.Poly3D.Length)
                                                {
                                                    dist2mn = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                                }
                                                Point1mn = _AGEN_mainform.Poly3D.GetPointAtDist(dist1);
                                                Point2mn = _AGEN_mainform.Poly3D.GetPointAtDist(dist2mn);
                                            }
                                            Polyline Rectangle_ML2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                                            Rectangle_ML2 = create_rectangle_Matchline(Point1mn, Point2mn, Colorindex);
                                            Rectangle_ML2.Layer = _AGEN_mainform.Layer_name_ML_rectangle;
                                            BTrecord.AppendEntity(Rectangle_ML2);
                                            Trans1.AddNewlyCreatedDBObject(Rectangle_ML2, true);

                                            string OBid1 = _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_handle].ToString();
                                            ObjectId obid = Functions.GetObjectId(ThisDrawing.Database, OBid1);
                                            if (obid != ObjectId.Null)
                                            {
                                                DBObject Dbobj1 = Trans1.GetObject(obid, OpenMode.ForWrite);
                                                Dbobj1.Erase();
                                            }


                                            #region USA
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_M1] = dist1;
                                            }
                                            #endregion

                                            #region CANADA
                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                            {
                                                double param1 = _AGEN_mainform.Poly3D.GetParameterAtPoint(Point1mn);
                                                double dist_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                                double b1 = -1.23456;
                                                double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1mn, dist_2d, _AGEN_mainform.dt_centerline, ref b1);
                                                if (sta_csf1 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_M1] = sta_csf1;
                                            }
                                            #endregion

                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_handle] = Rectangle_ML2.ObjectId.Handle.Value.ToString();
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_x] = (Rectangle_ML2.GetPoint3dAt(0).X + Rectangle_ML2.GetPoint3dAt(2).X) / 2;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_y] = (Rectangle_ML2.GetPoint3dAt(0).Y + Rectangle_ML2.GetPoint3dAt(2).Y) / 2;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle_ML2.GetPoint3dAt(1).X, Rectangle_ML2.GetPoint3dAt(1).Y, Rectangle_ML2.GetPoint3dAt(2).X, Rectangle_ML2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_Width] = Rectangle_ML2.GetPoint3dAt(1).DistanceTo(Rectangle_ML2.GetPoint3dAt(2));
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]][_AGEN_mainform.Col_Height] = Rectangle_ML2.GetPoint3dAt(0).DistanceTo(Rectangle_ML2.GetPoint3dAt(1));
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_Beg"] = Point1mn.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_Beg"] = Point1mn.Y;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["X_End"] = Point2mn.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[List2[i]]["Y_End"] = Point2mn.Y;

                                            Trans1.TransactionManager.QueueForGraphicsFlush();

                                        }
                                    }

                                }

                            }
                            #endregion
                        }
                        else
                        {
                            #region lista2.count=0

                            if (last_pt == new Point3d(1.234, 1.234, 1.234)) last_pt = _AGEN_mainform.Poly2D.StartPoint;

                            Point3d pt_on_line1 = last_pt;
                            double Param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_line1);
                            if (Param1 > _AGEN_mainform.Poly3D.EndParam) Param1 = _AGEN_mainform.Poly3D.EndParam;
                            Next_matchline = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param1);


                        l1234:
                            Point3d Point1m = new Point3d();
                            Point3d Point2m = new Point3d();

                            if (radioButton_3D_station.Checked == false)
                            {
                                #region 2D

                                last_pt = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);
                                zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                Trans1.TransactionManager.QueueForGraphicsFlush();

                                Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly2D, last_pt, 10);

                                if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                last_pt = _AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false);

                                if (_AGEN_mainform.Poly2D.EndPoint.DistanceTo(last_pt) < 0.0009)
                                {
                                    last_pt = _AGEN_mainform.Poly2D.EndPoint;
                                }

                                dist2 = _AGEN_mainform.Poly2D.GetDistAtPoint(last_pt);

                                if (Math.Round(dist1, 0) > Math.Round(dist2, 0))
                                {
                                    goto l1234;
                                }


                                Point1m = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);

                                Point2m = _AGEN_mainform.Poly2D.GetPointAtDist(dist2);
                                #endregion


                            }
                            else
                            {
                                bool recalc_last_pt = true;
                                if(dist1==0 && first_matchline_value>0)
                                {
                                    recalc_last_pt = false;
                                    dist1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(Param1);
                                }
                                if(recalc_last_pt==true) last_pt = _AGEN_mainform.Poly2D.GetPointAtDist(dist1);
                                zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                Trans1.TransactionManager.QueueForGraphicsFlush();


                                Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2();
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly3D, _AGEN_mainform.Poly2D, last_pt, 10);

                                if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                {
                                    Trans1.Commit();
                                    goto end1;
                                }

                                double p_l = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false));
                                if (Math.Round(p_l, 4) == Math.Round(_AGEN_mainform.Poly2D.EndParam, 4))
                                {
                                    p_l = _AGEN_mainform.Poly2D.EndParam;
                                }


                                last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(p_l);

                                dist2 = _AGEN_mainform.Poly3D.GetDistAtPoint(last_pt);

                                if (Math.Round(dist1, 0) > Math.Round(dist2, 0))
                                {
                                    goto l1234;
                                }


                                Point1m = _AGEN_mainform.Poly3D.GetPointAtDist(dist1);

                                Point2m = _AGEN_mainform.Poly3D.GetPointAtDist(dist2);
                            }



                            Polyline Rectangle_ML1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                            Rectangle_ML1 = create_rectangle_Matchline(Point1m, Point2m, Colorindex);
                            Rectangle_ML1.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                            BTrecord.AppendEntity(Rectangle_ML1);
                            Trans1.AddNewlyCreatedDBObject(Rectangle_ML1, true);


                            Line Line1 = new Line(Rectangle_ML1.GetPointAtParameter(2), Rectangle_ML1.GetPointAtParameter(3));
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                            Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                            Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                            Jig1.AddEntity(Rectangle_ML1);
                            Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                            if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Jig1.TransformEntities();
                            }

                            Trans1.TransactionManager.QueueForGraphicsFlush();


                            System.Data.DataRow Row1 = _AGEN_mainform.dt_sheet_index.NewRow();


                            Row1[_AGEN_mainform.Col_handle] = Rectangle_ML1.ObjectId.Handle.Value.ToString();
                            Row1[_AGEN_mainform.Col_x] = (Rectangle_ML1.GetPoint3dAt(0).X + Rectangle_ML1.GetPoint3dAt(2).X) / 2;
                            Row1[_AGEN_mainform.Col_y] = (Rectangle_ML1.GetPoint3dAt(0).Y + Rectangle_ML1.GetPoint3dAt(2).Y) / 2;
                            Row1[_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle_ML1.GetPoint3dAt(1).X, Rectangle_ML1.GetPoint3dAt(1).Y, Rectangle_ML1.GetPoint3dAt(2).X, Rectangle_ML1.GetPoint3dAt(2).Y) * 180 / Math.PI;
                            Row1[_AGEN_mainform.Col_Width] = Rectangle_ML1.GetPoint3dAt(1).DistanceTo(Rectangle_ML1.GetPoint3dAt(2));
                            Row1[_AGEN_mainform.Col_Height] = Rectangle_ML1.GetPoint3dAt(0).DistanceTo(Rectangle_ML1.GetPoint3dAt(1));
                            Row1["X_Beg"] = Point1m.X;
                            Row1["Y_Beg"] = Point1m.Y;
                            Row1["X_End"] = Point2m.X;
                            Row1["Y_End"] = Point2m.Y;

                            #region USA
                            if (_AGEN_mainform.COUNTRY == "USA")
                            {
                                Row1[_AGEN_mainform.Col_M1] = dist1;
                                Row1[_AGEN_mainform.Col_M2] = dist2;
                            }
                            #endregion

                            #region CANADA
                            if (_AGEN_mainform.COUNTRY == "CANADA")
                            {
                                double param1 = _AGEN_mainform.Poly3D.GetParameterAtPoint(Point1m);
                                double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                double b1 = -1.23456;
                                double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1m, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                if (sta_csf1 != -1.23456) Row1[_AGEN_mainform.Col_M1] = Math.Round(sta_csf1, _AGEN_mainform.round1);

                                double param2 = _AGEN_mainform.Poly3D.GetParameterAtPoint(Point2m);
                                double dist_2d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);

                                double b2 = -1.23456;
                                double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point2m, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                if (sta_csf2 != -1.23456) Row1[_AGEN_mainform.Col_M2] = Math.Round(sta_csf2, _AGEN_mainform.round1);
                            }
                            #endregion

                            _AGEN_mainform.dt_sheet_index.Rows.Add(Row1);



                            Colorindex = Colorindex + 1;
                            if (Colorindex > 7) Colorindex = 1;


                            dist1 = dist2;
                            #endregion

                        }

                        double Last_Matchline_Value = 0;

                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] != DBNull.Value &&
                                    _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] != DBNull.Value)
                            {
                                double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"]);
                                double y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"]);
                                Point3d Point_online2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                double param22 = _AGEN_mainform.Poly2D.GetParameterAtPoint(Point_online2);
                                if (param22 > _AGEN_mainform.Poly3D.EndParam) param22 = _AGEN_mainform.Poly3D.EndParam;
                                Last_Matchline_Value = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param22);

                                double div2 = 10;
                                if (_AGEN_mainform.round1 == 1) div2 = 100;
                                if (_AGEN_mainform.round1 == 2) div2 = 1000;
                                if (_AGEN_mainform.round1 == 3) div2 = 10000;

                                #region 2D
                                if (radioButton_3D_station.Checked == false)
                                {
                                    if (Math.Round(_AGEN_mainform.Poly2D.Length, 0) > Math.Round(Last_Matchline_Value, 0))
                                    {
                                        if (Last_Matchline_Value >= _AGEN_mainform.Poly2D.Length)
                                        {
                                            Last_Matchline_Value = Math.Floor(Math.Round(_AGEN_mainform.Poly2D.Length * div2, _AGEN_mainform.round1 + 1)) / div2;
                                        }

                                        last_pt = _AGEN_mainform.Poly2D.GetPointAtDist(Last_Matchline_Value);

                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                        zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                    l1235:
                                        Point3d Point2m = new Point3d();
                                        Point3d Point1m = new Point3d();

                                        Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along2D_manual_pt2();
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly2D, last_pt, 10);

                                        if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                        {
                                            Trans1.Commit();
                                            goto end1;
                                        }

                                        last_pt = _AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false);

                                        double Lastdist2 = _AGEN_mainform.Poly2D.GetDistAtPoint(last_pt);

                                        if (Math.Round(Last_Matchline_Value, 0) > Math.Round(Lastdist2, 0))
                                        {
                                            goto l1235;
                                        }


                                        Point1m = _AGEN_mainform.Poly2D.GetPointAtDist(Last_Matchline_Value);

                                        Point2m = _AGEN_mainform.Poly2D.GetPointAtDist(Lastdist2);

                                        Polyline Rectangle_ml3 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                        Rectangle_ml3 = create_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                        Rectangle_ml3.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                                        BTrecord.AppendEntity(Rectangle_ml3);
                                        Trans1.AddNewlyCreatedDBObject(Rectangle_ml3, true);


                                        Line Line1 = new Line(Rectangle_ml3.GetPointAtParameter(2), Rectangle_ml3.GetPointAtParameter(3));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                        Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                        Jig1.AddEntity(Rectangle_ml3);
                                        Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                        if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Jig1.TransformEntities();
                                        }

                                        Trans1.TransactionManager.QueueForGraphicsFlush();




                                        _AGEN_mainform.dt_sheet_index.Rows.Add();
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle_ml3.ObjectId.Handle.Value.ToString();
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle_ml3.GetPoint3dAt(0).X + Rectangle_ml3.GetPoint3dAt(2).X) / 2;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle_ml3.GetPoint3dAt(0).Y + Rectangle_ml3.GetPoint3dAt(2).Y) / 2;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle_ml3.GetPoint3dAt(1).X, Rectangle_ml3.GetPoint3dAt(1).Y, Rectangle_ml3.GetPoint3dAt(2).X, Rectangle_ml3.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle_ml3.GetPoint3dAt(1).DistanceTo(Rectangle_ml3.GetPoint3dAt(2));
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle_ml3.GetPoint3dAt(0).DistanceTo(Rectangle_ml3.GetPoint3dAt(1));
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1m.X;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1m.Y;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2m.X;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2m.Y;
                                        #region USA
                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = Last_Matchline_Value;
                                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = Lastdist2;
                                        }
                                        #endregion





                                        Colorindex = Colorindex + 1;
                                        if (Colorindex > 7) Colorindex = 1;
                                        if (Math.Round(Lastdist2, 0) == Math.Round(_AGEN_mainform.Poly2D.Length, 0))
                                        {
                                            Lastdist2 = _AGEN_mainform.Poly2D.Length;
                                        }


                                        Last_Matchline_Value = Lastdist2;

                                        if (Math.Round(Lastdist2, 0) < Math.Round(_AGEN_mainform.Poly2D.Length, 0))
                                        {
                                            goto l1235;
                                        }



                                    }
                                }

                                #endregion


                                if (radioButton_3D_station.Checked == true)
                                {

                                    if (Math.Round(_AGEN_mainform.Poly3D.Length, 0) > Math.Round(Last_Matchline_Value, 0))
                                    {

                                        if (Last_Matchline_Value >= _AGEN_mainform.Poly3D.Length)
                                        {
                                            Last_Matchline_Value = Math.Floor(Math.Round(_AGEN_mainform.Poly3D.Length * div2, _AGEN_mainform.round1 + 1)) / div2;
                                        }



                                        last_pt = _AGEN_mainform.Poly3D.GetPointAtDist(Last_Matchline_Value);

                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                        zoom_to_Point(last_pt, _AGEN_mainform.Vw_height * 1.5);
                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                    l12353:
                                        Point3d Point2m = new Point3d();
                                        Point3d Point1m = new Point3d();

                                        Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2 Jig2m = new Alignment_mdi.Jig_rectangle_viewport_along3D_manual_pt2();
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2m = Jig2m.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Match_distance, _AGEN_mainform.Vw_height, _AGEN_mainform.Poly3D, _AGEN_mainform.Poly2D, last_pt, 10);

                                        if (Result_point_m2m.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                                        {
                                            Trans1.Commit();
                                            goto end1;
                                        }

                                        double p_l = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(Result_point_m2m.Value, Vector3d.ZAxis, false));
                                        if (Math.Round(p_l, 4) == Math.Round(_AGEN_mainform.Poly2D.EndParam, 4))
                                        {
                                            p_l = _AGEN_mainform.Poly2D.EndParam;
                                        }


                                        last_pt = _AGEN_mainform.Poly3D.GetPointAtParameter(p_l);


                                        double Lastdist2 = _AGEN_mainform.Poly3D.GetDistAtPoint(last_pt);
                                        if (Math.Abs(Lastdist2 - _AGEN_mainform.Poly3D.Length) < 0.1) Lastdist2 = _AGEN_mainform.Poly3D.Length - 0.001;

                                        if (Math.Round(Last_Matchline_Value, 0) > Math.Round(Lastdist2, 0))
                                        {
                                            goto l12353;
                                        }


                                        Point1m = _AGEN_mainform.Poly3D.GetPointAtDist(Last_Matchline_Value);

                                        Point2m = _AGEN_mainform.Poly3D.GetPointAtDist(Lastdist2);

                                        Polyline Rectangle_ml3 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                        Rectangle_ml3 = create_rectangle_Matchline(Point1m, Point2m, Colorindex);
                                        Rectangle_ml3.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                                        BTrecord.AppendEntity(Rectangle_ml3);
                                        Trans1.AddNewlyCreatedDBObject(Rectangle_ml3, true);


                                        Line Line1 = new Line(Rectangle_ml3.GetPointAtParameter(2), Rectangle_ml3.GetPointAtParameter(3));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.StartPoint));
                                        Line1.TransformBy(Matrix3d.Scaling(10000, Line1.EndPoint));

                                        Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Point2m, Line1);
                                        Jig1.AddEntity(Rectangle_ml3);
                                        Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                        if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Jig1.TransformEntities();
                                        }

                                        Trans1.TransactionManager.QueueForGraphicsFlush();




                                        _AGEN_mainform.dt_sheet_index.Rows.Add();

                                        #region USA
                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = Last_Matchline_Value;
                                            _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = Lastdist2;
                                        }
                                        #endregion

                                        #region CANADA
                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                        {

                                            double param1 = _AGEN_mainform.Poly3D.GetParameterAtDistance(Last_Matchline_Value);
                                            double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);

                                            double b1 = -1.23456;
                                            double sta_csf1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point1m, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                            if (sta_csf1 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = sta_csf1;

                                            double param2 = _AGEN_mainform.Poly3D.GetParameterAtDistance(Lastdist2);
                                            double dist_2d2 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);

                                            double b2 = -1.23456;
                                            double sta_csf2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, Point2m, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                            if (sta_csf2 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = sta_csf2;
                                        }
                                        #endregion


                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle_ml3.ObjectId.Handle.Value.ToString();
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle_ml3.GetPoint3dAt(0).X + Rectangle_ml3.GetPoint3dAt(2).X) / 2;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle_ml3.GetPoint3dAt(0).Y + Rectangle_ml3.GetPoint3dAt(2).Y) / 2;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle_ml3.GetPoint3dAt(1).X, Rectangle_ml3.GetPoint3dAt(1).Y, Rectangle_ml3.GetPoint3dAt(2).X, Rectangle_ml3.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle_ml3.GetPoint3dAt(1).DistanceTo(Rectangle_ml3.GetPoint3dAt(2));
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle_ml3.GetPoint3dAt(0).DistanceTo(Rectangle_ml3.GetPoint3dAt(1));
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1m.X;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1m.Y;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2m.X;
                                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2m.Y;




                                        Colorindex = Colorindex + 1;
                                        if (Colorindex > 7) Colorindex = 1;
                                        if (Math.Round(Lastdist2, 0) == Math.Round(_AGEN_mainform.Poly3D.Length, 0))
                                        {
                                            Lastdist2 = _AGEN_mainform.Poly3D.Length;
                                        }


                                        Last_Matchline_Value = Lastdist2;

                                        if (Math.Round(Lastdist2, 0) < Math.Round(_AGEN_mainform.Poly3D.Length, 0))
                                        {
                                            goto l12353;
                                        }



                                    }
                                }
                            }
                        }
                        Editor1.WriteMessage("\nCommand:");

                        Trans1.Commit();



                    }
                }


            end1:

                if (_AGEN_mainform.dt_sheet_index != null)
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        Populate_data_table_matchline_file_names();
                        dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                        dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                        dataGridView_sheet_index.EnableHeadersVisualStyles = false;

                        Append_ML_object_data();


                        string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjF.Length > 0)
                        {
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }
                        }


                        string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;




                        round_sheet_index_data_table(poly_length);

                        delete_centerlines();
                        label_not_saved.Visible = true;

                    }
                }
            }
            catch (System.Exception ex)
            {
                set_enable_true();
                delete_centerlines();
                MessageBox.Show(ex.Message);
            }

            set_enable_true();


            Ag.WindowState = FormWindowState.Normal;
        }


        private void button_adjust_rectangle_Click(object sender, EventArgs e)
        {
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


            if (_AGEN_mainform.dt_centerline == null)
            {
                MessageBox.Show("you do not have picked the centerline\r\noperation aborted");


                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();

                return;
            }

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("the centerline file d\r\noperation aborted");


                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();
                return;
            }

            if (_AGEN_mainform.dt_sheet_index == null)
            {
                MessageBox.Show("the sheet index table is null\r\noperation aborted");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("No data for sheet indexes found\r\noperation aborted");
                return;
            }


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)

                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
                Ag.WindowState = FormWindowState.Minimized;
            }

            set_enable_false();
            Erase_viewports_templates();
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                    delete_centerlines();
                    lista_del.Add(_AGEN_mainform.Poly3D.ObjectId);
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        string Path_toCL = "";
                        string Fisier_si = "";
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            Path_toCL = ProjF + _AGEN_mainform.cl_excel_name;

                            if (System.IO.File.Exists(Path_toCL) == false)
                            {
                                set_enable_true();
                                MessageBox.Show("No centerline file loaded");
                                return;
                            }

                            Fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                            if (System.IO.File.Exists(Fisier_si) == false)
                            {
                                set_enable_true();
                                MessageBox.Show("No sheet index file loaded");
                                return;
                            }
                        }
                        else
                        {
                            set_enable_true();
                            MessageBox.Show("the project folder does not exist");
                            return;
                        }


                        _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);



                        if (_AGEN_mainform.Poly2D == null && radioButton_3D_station.Checked == true)
                        {
                            _AGEN_mainform.Poly3D = Trans1.GetObject(_AGEN_mainform.Poly3D.ObjectId, OpenMode.ForRead) as Polyline3d;

                            if (_AGEN_mainform.Poly3D == null)
                            {
                                set_enable_true();
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                MessageBox.Show("there is no centerline into the current drawing");
                                return;

                            }
                            _AGEN_mainform.Poly2D = Functions.Build_2dpoly_from_3d(_AGEN_mainform.Poly3D);
                        }
                        else if (_AGEN_mainform.Poly2D == null && radioButton_3D_station.Checked == false)
                        {
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            MessageBox.Show("there is no centerline into the current drawing");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsrec = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect rectangle:");
                        Prompt_optionsrec.SetRejectMessage("\nYou did not selected a polyline");
                        Prompt_optionsrec.AddAllowedClass(typeof(Polyline), true);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_rec = Editor1.GetEntity(Prompt_optionsrec);
                        if (Rezultat_rec.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            return;
                        }

                        Polyline Rect_0 = (Polyline)Trans1.GetObject(Rezultat_rec.ObjectId, OpenMode.ForWrite);

                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        //Dim BTrecord_MS As Autodesk.AutoCAD.DatabaseServices.BlockTableRecord = Trans1.GetObject(BlockTable1(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        ObjectId Obj_id_old = Rect_0.ObjectId;



                        string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();
                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                        }
                        else
                        {
                            if (Scale1.Contains(":") == true)
                            {
                                Scale1 = Scale1.Replace("1:", "");
                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                                }
                            }
                            else
                            {
                                string inch = "\u0022";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                                }

                                inch = "\u0094";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                                }

                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                                }
                            }
                        }

                        if (_AGEN_mainform.dt_sheet_index != null)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {
                                int Index0 = -1;
                                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {
                                    if (Obj_id_old == Functions.GetObjectId(ThisDrawing.Database, (string)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_handle]))
                                    {
                                        Index0 = i;
                                        i = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                    }
                                }

                                if (Index0 != -1)
                                {

                                    Polyline poly1 = new Polyline();
                                    poly1.AddVertexAt(0, Rect_0.GetPoint2dAt(2), 0, 0, 0);
                                    poly1.AddVertexAt(1, Rect_0.GetPoint2dAt(3), 0, 0, 0);

                                    poly1.TransformBy(Matrix3d.Scaling(10000, new Point3d((poly1.StartPoint.X + poly1.EndPoint.X) / 2,
                                                                                            (poly1.StartPoint.Y + poly1.EndPoint.Y) / 2,
                                                                                                _AGEN_mainform.Poly2D.Elevation)));



                                    if (_AGEN_mainform.dt_sheet_index.Rows[Index0]["X_Beg"] != DBNull.Value &&
                                        _AGEN_mainform.dt_sheet_index.Rows[Index0]["Y_Beg"] != DBNull.Value &&
                                        _AGEN_mainform.dt_sheet_index.Rows[Index0]["X_End"] != DBNull.Value &&
                                        _AGEN_mainform.dt_sheet_index.Rows[Index0]["Y_End"] != DBNull.Value)
                                    {
                                        double x1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[Index0]["X_Beg"]);
                                        double y1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[Index0]["Y_Beg"]);
                                        double x2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[Index0]["X_End"]);
                                        double y2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[Index0]["Y_End"]);

                                        Point3d Point_on2d1 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x1, y1, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                        Point3d Point_on2d2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                        double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(Point_on2d1);
                                        double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(Point_on2d2);

                                        if (param2 > _AGEN_mainform.Poly3D.EndParam) param2 = _AGEN_mainform.Poly3D.EndParam;

                                        double M2 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param2);

                                        Point3dCollection Col1 = new Point3dCollection();
                                        if (radioButton_3D_station.Checked == false)
                                        {
                                            Col1 = Functions.Intersect_with_extend(poly1, _AGEN_mainform.Poly2D);
                                        }
                                        else
                                        {
                                            Col1 = Functions.Intersect_with_extend_2d_3d(_AGEN_mainform.Poly3D, poly1);
                                        }
                                        if (Col1.Count == 0)
                                        {
                                            MessageBox.Show("The rectangle does not intersect the centerline....");
                                            set_enable_true();
                                            return;
                                        }

                                        for (int i = 0; i < Col1.Count; ++i)
                                        {
                                            try
                                            {

                                                if (radioButton_3D_station.Checked == false)
                                                {
                                                    if (Math.Round(_AGEN_mainform.Poly2D.GetDistAtPoint(Col1[i]), 0) == M2)
                                                    {
                                                        M2 = _AGEN_mainform.Poly2D.GetDistAtPoint(Col1[i]);
                                                        i = Col1.Count;
                                                    }
                                                }
                                                else
                                                {
                                                    double Param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(Col1[i]);
                                                    if (Math.Round(Param1, 4) == Math.Round(_AGEN_mainform.Poly2D.EndParam, 4))
                                                    {
                                                        Param1 = _AGEN_mainform.Poly2D.EndParam;
                                                    }
                                                    double dist1 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param1);
                                                    if (Math.Round(dist1, 0) == M2)
                                                    {
                                                        M2 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(Param1);
                                                        i = Col1.Count;
                                                    }
                                                }
                                            }
                                            catch (System.Exception ex)
                                            {
                                                MessageBox.Show("The rectangle does not intersect the centerline....");
                                                set_enable_true();
                                                return;
                                            }

                                        }
                                        Point3d Ptt = new Point3d();
                                        if (radioButton_3D_station.Checked == false)
                                        {
                                            if (Math.Round(_AGEN_mainform.Poly2D.Length, 1) == Math.Round(M2, 1)) M2 = _AGEN_mainform.Poly2D.Length;
                                            Ptt = _AGEN_mainform.Poly2D.GetPointAtDist(M2);
                                        }
                                        else
                                        {
                                            if (Math.Abs(_AGEN_mainform.Poly3D.Length - M2) < 0.1) M2 = _AGEN_mainform.Poly3D.Length - 0.001;
                                            Ptt = _AGEN_mainform.Poly3D.GetPointAtDist(M2);
                                            Ptt = new Point3d(Ptt.X, Ptt.Y, 0);

                                        }

                                        Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down Jig1 = new Alignment_mdi.Jig_rectangle_viewport_SHEET_CUTTER_manual_up_down(Ptt, new Line(poly1.StartPoint, poly1.EndPoint));
                                        Jig1.AddEntity(Rect_0);
                                        Autodesk.AutoCAD.EditorInput.PromptResult jigRes = ThisDrawing.Editor.Drag(Jig1);
                                        if (jigRes.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Jig1.TransformEntities();
                                            _AGEN_mainform.dt_sheet_index.Rows[Index0][_AGEN_mainform.Col_x] = (Rect_0.GetPoint3dAt(0).X + Rect_0.GetPoint3dAt(2).X) / 2;
                                            _AGEN_mainform.dt_sheet_index.Rows[Index0][_AGEN_mainform.Col_y] = (Rect_0.GetPoint3dAt(0).Y + Rect_0.GetPoint3dAt(2).Y) / 2;
                                        }

                                        Trans1.TransactionManager.QueueForGraphicsFlush();

                                    }
                                }
                            }
                        }

                        Trans1.Commit();
                        Editor1.WriteMessage("\nCommand:");
                    }

                }

                if (_AGEN_mainform.dt_sheet_index != null && _AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    Append_ML_object_data();
                    dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                    dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                    dataGridView_sheet_index.EnableHeadersVisualStyles = false;
                    label_not_saved.Visible = true;

                }

            }
            catch (System.Exception ex)
            {
                set_enable_true();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }


        private void button_recover_matclines_Click(object sender, EventArgs e)
        {

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }

            set_enable_false();
            _AGEN_mainform.tpage_processing.Show();
            if (_AGEN_mainform.dt_sheet_index != null && _AGEN_mainform.dt_centerline != null)
            {
                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0 && _AGEN_mainform.dt_centerline.Rows.Count > 0)
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
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                delete_polyline_with_OD(_AGEN_mainform.Layer_name_ML_rectangle, "Agen_SheetIndex_ML");
                                delete_mtext_with_OD(_AGEN_mainform.Layer_name_ML_rectangle, "Agen_SheetIndex_ML");


                                Polyline poly2d = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                Polyline3d poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                if (radioButton_3D_station.Checked == false)
                                {
                                    BTrecord.AppendEntity(poly2d);
                                    Trans1.AddNewlyCreatedDBObject(poly2d, true);
                                }

                                #region USA - add MEASURED dt_station eq
                                if (_AGEN_mainform.COUNTRY == "USA" && _AGEN_mainform.dt_station_equation != null)
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
                                                double param1 = poly2d.GetParameterAtPoint(pt_on_2d);

                                                if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                                                double eq_meas = poly3d.GetDistanceAtParameter(param1);
                                                _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                            }
                                        }
                                    }

                                }
                                #endregion



                                int CI = 1;
                                Functions.Creaza_layer(_AGEN_mainform.Layer_name_ML_rectangle, 4, false);
                                Polyline Rectanglep = null;

                                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {
                                    string col_M1 = "StaBeg";
                                    string col_eq1 = "Disp_StaBeg";
                                    string col_M2 = "StaEnd";
                                    string col_eq2 = "Disp_StaEnd";


                                    double Cx = 0;
                                    double Cy = 0;
                                    double rotation = 0;
                                    double width1 = 0;
                                    double height1 = 0;

                                    double Cxn = 0;
                                    double Cyn = 0;
                                    double rotationn = 0;
                                    double width1n = 0;
                                    double height1n = 0;

                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value)
                                    {
                                        Cx = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x];
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle X value for sheet index in row " + (i).ToString());
                                        return;
                                    }
                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                    {
                                        Cy = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y];
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Y value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot] != DBNull.Value)
                                    {
                                        rotation = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot] * Math.PI / 180;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle ROTATION value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Height] != DBNull.Value)
                                    {
                                        height1 = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Height];
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Height value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Width] != DBNull.Value)
                                    {
                                        width1 = (double)_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Width];
                                    }
                                    else
                                    {
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Width value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (i < _AGEN_mainform.dt_sheet_index.Rows.Count - 1)
                                    {


                                        if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_x] != DBNull.Value)
                                        {
                                            Cxn = (double)_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_x];
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle X value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }
                                        if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_y] != DBNull.Value)
                                        {
                                            Cyn = (double)_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_y];
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Y value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_rot] != DBNull.Value)
                                        {
                                            rotationn = (double)_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_rot] * Math.PI / 180;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle ROTATION value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_Height] != DBNull.Value)
                                        {
                                            height1n = (double)_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_Height];
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Height value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_Width] != DBNull.Value)
                                        {
                                            width1n = (double)_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_Width];
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Width value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }
                                    }





                                    Polyline Rectangle1 = creaza_rectangle_from_one_point(new Point3d(Cx, Cy, 0), rotation, width1, height1, CI);
                                    Rectangle1.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                                    Polyline p1 = new Polyline();
                                    p1.AddVertexAt(0, Rectangle1.GetPoint2dAt(0), 0, 0, 0);
                                    p1.AddVertexAt(1, Rectangle1.GetPoint2dAt(1), 0, 0, 0);
                                    p1.Elevation = poly2d.Elevation;

                                    Polyline p2 = new Polyline();
                                    p2.AddVertexAt(0, Rectangle1.GetPoint2dAt(2), 0, 0, 0);
                                    p2.AddVertexAt(1, Rectangle1.GetPoint2dAt(3), 0, 0, 0);
                                    p2.Elevation = poly2d.Elevation;

                                    Point3dCollection colm1 = Functions.Intersect_on_both_operands(p1, poly2d);
                                    Point3dCollection colm2 = Functions.Intersect_on_both_operands(p2, poly2d);


                                    if (colm1.Count > 0)
                                    {
                                        if (colm1.Count == 1)
                                        {

                                            Point3d Point_online1 = poly2d.GetClosestPointTo(new Point3d(colm1[0].X, colm1[0].Y, poly2d.Elevation), Vector3d.ZAxis, false);

                                            _AGEN_mainform.dt_sheet_index.Rows[i]["X_Beg"] = Point_online1.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[i]["Y_Beg"] = Point_online1.Y;

                                            Point3d pt_on_poly1 = poly2d.GetClosestPointTo(Point_online1, Vector3d.ZAxis, false);

                                            double parameter1 = poly2d.GetParameterAtPoint(pt_on_poly1);

                                            if (parameter1 > poly3d.EndParam) parameter1 = poly3d.EndParam;
                                            double mdist1 = Math.Round(poly3d.GetDistanceAtParameter(parameter1), _AGEN_mainform.round1);
                                            #region USA
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = mdist1;
                                            }
                                            #endregion
                                            #region CANADA
                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                            {

                                                double dist_2d1 = poly2d.GetDistanceAtParameter(parameter1);
                                                double b1 = -1.23456;
                                                double sta_csf1 = Functions.get_stationCSF_from_point(poly2d, pt_on_poly1, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                                if (sta_csf1 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = Math.Round(sta_csf1, _AGEN_mainform.round1);

                                                _AGEN_mainform.dt_sheet_index.Rows[i][5] = DBNull.Value;
                                            }
                                            #endregion

                                        }
                                        else
                                        {
                                            if (i == 0)
                                            {
                                                for (int j = 0; j < colm1.Count; ++j)
                                                {
                                                    if (poly2d.StartPoint.DistanceTo(colm1[j]) < 0.01)
                                                    {
                                                        _AGEN_mainform.dt_sheet_index.Rows[i]["X_Beg"] = poly2d.StartPoint.X;
                                                        _AGEN_mainform.dt_sheet_index.Rows[i]["Y_Beg"] = poly2d.StartPoint.Y;
                                                        _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = 0;
                                                        j = colm1.Count;

                                                    }
                                                }
                                            }
                                            else
                                            {
                                                Polyline p2p = new Polyline();
                                                p2p.AddVertexAt(0, Rectanglep.GetPoint2dAt(2), 0, 0, 0);
                                                p2p.AddVertexAt(1, Rectanglep.GetPoint2dAt(3), 0, 0, 0);
                                                p2p.Elevation = poly2d.Elevation;
                                                Point3dCollection colm1p = Functions.Intersect_on_both_operands(p1, p2p);

                                                if (colm1p.Count == 1)
                                                {
                                                    for (int j = 0; j < colm1.Count; ++j)
                                                    {
                                                        if (colm1p[0].DistanceTo(colm1[j]) < 0.01)
                                                        {

                                                            Point3d Point_online1 = poly2d.GetClosestPointTo(new Point3d(colm1[j].X, colm1[j].Y, poly2d.Elevation), Vector3d.ZAxis, false);

                                                            _AGEN_mainform.dt_sheet_index.Rows[i]["X_Beg"] = Point_online1.X;
                                                            _AGEN_mainform.dt_sheet_index.Rows[i]["Y_Beg"] = Point_online1.Y;

                                                            double parameter1 = poly2d.GetParameterAtPoint(Point_online1);
                                                            if (parameter1 > poly3d.EndParam) parameter1 = poly3d.EndParam;

                                                            double mdist1 = Math.Round(poly3d.GetDistanceAtParameter(parameter1), _AGEN_mainform.round1);
                                                            #region USA
                                                            if (_AGEN_mainform.COUNTRY == "USA")
                                                            {
                                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = mdist1;
                                                            }
                                                            #endregion
                                                            #region CANADA
                                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                                            {

                                                                double dist_2d1 = poly2d.GetDistanceAtParameter(parameter1);
                                                                double b1 = -1.23456;
                                                                double sta_csf1 = Functions.get_stationCSF_from_point(poly2d, Point_online1, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                                                if (sta_csf1 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = Math.Round(sta_csf1, _AGEN_mainform.round1);

                                                                _AGEN_mainform.dt_sheet_index.Rows[i][5] = DBNull.Value;
                                                            }
                                                            #endregion

                                                            j = colm1.Count;

                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = -1;

                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        _AGEN_mainform.dt_sheet_index.Rows[i][col_M1] = -1;

                                    }


                                    if (colm2.Count > 0)
                                    {
                                        if (colm2.Count == 1)
                                        {
                                            Point3d Point_online2 = poly2d.GetClosestPointTo(new Point3d(colm2[0].X, colm2[0].Y, poly2d.Elevation), Vector3d.ZAxis, false);

                                            _AGEN_mainform.dt_sheet_index.Rows[i]["X_End"] = Point_online2.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[i]["Y_End"] = Point_online2.Y;

                                            Point3d pt_on_poly2 = poly2d.GetClosestPointTo(Point_online2, Vector3d.ZAxis, false);

                                            double parameter2 = poly2d.GetParameterAtPoint(pt_on_poly2);

                                            if (parameter2 > poly3d.EndParam) parameter2 = poly3d.EndParam;


                                            double mdist2 = Math.Round(poly3d.GetDistanceAtParameter(parameter2), _AGEN_mainform.round1);
                                            #region USA
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = mdist2;
                                            }
                                            #endregion
                                            #region CANADA
                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                            {

                                                double dist_2d2 = poly2d.GetDistanceAtParameter(parameter2);
                                                double b2 = -1.23456;
                                                double sta_csf2 = Functions.get_stationCSF_from_point(poly2d, pt_on_poly2, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                                if (sta_csf2 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = Math.Round(sta_csf2, _AGEN_mainform.round1);

                                                _AGEN_mainform.dt_sheet_index.Rows[i][6] = DBNull.Value;
                                            }
                                            #endregion

                                        }
                                        else
                                        {


                                            if (i == _AGEN_mainform.dt_sheet_index.Rows.Count - 1)
                                            {

                                                for (int j = 0; j < colm2.Count; ++j)
                                                {
                                                    if (poly2d.StartPoint.DistanceTo(colm2[j]) < 0.01)
                                                    {
                                                        _AGEN_mainform.dt_sheet_index.Rows[i]["X_End"] = poly2d.EndPoint.X;
                                                        _AGEN_mainform.dt_sheet_index.Rows[i]["Y_End"] = poly2d.EndPoint.Y;
                                                        double lenn1 = poly3d.Length;
                                                        double div1 = 10;
                                                        if (_AGEN_mainform.round1 == 1) div1 = 100;
                                                        if (_AGEN_mainform.round1 == 2) div1 = 1000;
                                                        if (_AGEN_mainform.round1 == 3) div1 = 10000;
                                                        lenn1 = Math.Floor(Math.Round(poly3d.Length * div1, _AGEN_mainform.round1 + 1)) / div1;

                                                        if (_AGEN_mainform.COUNTRY == "USA")
                                                        {
                                                            _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = lenn1;
                                                        }

                                                        #region CANADA
                                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                                        {
                                                            _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = _AGEN_mainform.dt_centerline.Rows[_AGEN_mainform.dt_centerline.Rows.Count - 1][_AGEN_mainform.Col_3DSta];
                                                        }
                                                        #endregion
                                                        j = colm2.Count;

                                                    }
                                                }


                                            }
                                            else
                                            {
                                                Polyline Rectanglen = creaza_rectangle_from_one_point(new Point3d(Cxn, Cyn, 0), rotationn, width1n, height1n, CI);
                                                Polyline p2n = new Polyline();
                                                p2n.AddVertexAt(0, Rectanglen.GetPoint2dAt(0), 0, 0, 0);
                                                p2n.AddVertexAt(1, Rectanglen.GetPoint2dAt(1), 0, 0, 0);
                                                p2n.Elevation = poly2d.Elevation;
                                                Point3dCollection colm1n = Functions.Intersect_on_both_operands(p2, p2n);

                                                if (colm1n.Count == 1)
                                                {
                                                    for (int j = 0; j < colm2.Count; ++j)
                                                    {
                                                        if (colm1n[0].DistanceTo(colm2[j]) < 0.01)
                                                        {


                                                            Point3d Point_online2 = poly2d.GetClosestPointTo(new Point3d(colm2[j].X, colm2[j].Y, poly2d.Elevation), Vector3d.ZAxis, false);

                                                            _AGEN_mainform.dt_sheet_index.Rows[i]["X_End"] = Point_online2.X;
                                                            _AGEN_mainform.dt_sheet_index.Rows[i]["Y_End"] = Point_online2.Y;

                                                            double parameter2 = poly2d.GetParameterAtPoint(Point_online2);
                                                            double mdist2 = Math.Round(poly3d.GetDistanceAtParameter(parameter2), _AGEN_mainform.round1, MidpointRounding.ToEven);
                                                            #region USA
                                                            if (_AGEN_mainform.COUNTRY == "USA")
                                                            {
                                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = mdist2;
                                                            }
                                                            #endregion
                                                            #region CANADA
                                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                                            {

                                                                double dist_2d2 = poly2d.GetDistanceAtParameter(parameter2);
                                                                double b2 = -1.23456;
                                                                double sta_csf2 = Functions.get_stationCSF_from_point(poly2d, Point_online2, dist_2d2, _AGEN_mainform.dt_centerline, ref b2);
                                                                if (sta_csf2 != -1.23456) _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = Math.Round(sta_csf2, _AGEN_mainform.round1);

                                                                _AGEN_mainform.dt_sheet_index.Rows[i][6] = DBNull.Value;
                                                            }
                                                            #endregion

                                                            j = colm2.Count;





                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = -1;

                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        _AGEN_mainform.dt_sheet_index.Rows[i][col_M2] = -1;

                                    }



                                    #region USA - populate sta eq columns
                                    if (_AGEN_mainform.COUNTRY == "USA")
                                    {
                                        if (_AGEN_mainform.dt_sheet_index.Rows[i][col_M1] != DBNull.Value && Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][col_M1]) >= 0)
                                        {
                                            double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][col_M1]);
                                            if (_AGEN_mainform.dt_station_equation != null)
                                            {
                                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_eq1] = Math.Round(Functions.Station_equation_ofV2(M1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_eq1] = DBNull.Value;
                                                }
                                            }
                                            else
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_eq1] = DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[i][col_eq1] = DBNull.Value;
                                        }

                                        if (_AGEN_mainform.dt_sheet_index.Rows[i][col_M2] != DBNull.Value && Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][col_M2]) >= 0)
                                        {
                                            double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][col_M2]);
                                            if (_AGEN_mainform.dt_station_equation != null)
                                            {
                                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_eq2] = Math.Round(Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows[i][col_eq2] = DBNull.Value;
                                                }
                                            }
                                            else
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows[i][col_eq2] = DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            _AGEN_mainform.dt_sheet_index.Rows[i][col_eq2] = DBNull.Value;
                                        }
                                    }
                                    #endregion




                                    BTrecord.AppendEntity(Rectangle1);
                                    Trans1.AddNewlyCreatedDBObject(Rectangle1, true);

                                    _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_handle] = Rectangle1.ObjectId.Handle.Value.ToString();
                                    CI = CI + 1;
                                    if (CI > 7) CI = 1;



                                    Rectanglep = Rectangle1;
                                }





                                Create_ML_object_data();
                                Append_ML_object_data();

                                if (radioButton_3D_station.Checked == false)
                                {
                                    poly3d.Erase();
                                }

                                Trans1.Commit();
                                dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;


                                label_not_saved.Visible = true;




                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
            }

        }





        private void Erase_matchlines_templates()
        {
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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            _AGEN_mainform.tpage_setup.delete_entities_with_OD(_AGEN_mainform.Layer_name_ML_rectangle, "Agen_SheetIndex_ML");

                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

        }

        private void Erase_Ms_blocks()
        {

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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            foreach (ObjectId Odid in BTrecord)
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(Odid, OpenMode.ForRead);
                                if (Ent1 != null)
                                {
                                    if (Ent1.Layer == _AGEN_mainform.Layer_even || Ent1.Layer == _AGEN_mainform.Layer_odd)
                                    {
                                        Ent1.UpgradeOpen();
                                        Ent1.Erase();
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

        }

        public bool get_radioButton_use3D_stations()
        {
            return radioButton_3D_station.Checked;
        }
        public void set_radioButton_use3D_stations(bool value)
        {
            radioButton_3D_station.Checked = value;
        }

        public void set_radioButton_use2D_stations(bool value)
        {
            radioButton_2D_station.Checked = value;
        }

        private void TextBox_keypress_only_doubles(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }

        private void TextBox_keypress_only_integers(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_integer_pozitive_at_keypress(sender, e);
        }



        private void button_reload_sheet_index_Click(object sender, EventArgs e)
        {
            set_enable_false();
            _AGEN_mainform.dt_sheet_index = Functions.Build_Data_table_sheet_index_from_object_data();
            dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.EnableHeadersVisualStyles = false;
            set_enable_true();


        }

        private void button_insert_na_ms_Click(object sender, EventArgs e)
        {

            if (_AGEN_mainform.dt_sheet_index == null)
            {
                MessageBox.Show("no sheet index information loaded");
                return;
            }
            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("no sheet index information loaded");
                return;
            }

            _AGEN_mainform.tpage_processing.Show();
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
                        if (BlockTable1.Has(_AGEN_mainform.NorthArrowMS) == true)
                        {

                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            double h1 = _AGEN_mainform.Vw_height;
                            double w1 = _AGEN_mainform.Vw_width;
                            double xps = _AGEN_mainform.Vw_ps_x;
                            double yps = _AGEN_mainform.Vw_ps_y;
                            double scale1 = _AGEN_mainform.Vw_scale;
                            double xna = _AGEN_mainform.NA_x;
                            double yna = _AGEN_mainform.NA_y;
                            int lr = 1;
                            if (_AGEN_mainform.Left_to_Right == false) lr = -1;


                            if (h1 <= 0 || w1 <= 0 || scale1 <= 0)
                            {
                                _AGEN_mainform.tpage_processing.Hide();
                                MessageBox.Show("please set up main viewport scale, width and height");
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                return;
                            }


                            if (xna <= 0 || yna <= 0)
                            {
                                _AGEN_mainform.tpage_processing.Hide();
                                MessageBox.Show("please set up north arrow position in paper space");
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                return;
                            }

                            #region object data stationing
                            Erase_northarrows();
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            Functions.Create_northarrow_od_table();

                            List<object> Lista_val = new List<object>();
                            List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                            if (segment1 == "not defined") segment1 = "";
                            Lista_val.Add(segment1);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            #endregion


                            Functions.Creaza_layer(_AGEN_mainform.Layer_even, 7, true);
                            Functions.Creaza_layer(_AGEN_mainform.Layer_odd, 7, true);

                            h1 = h1 / scale1;
                            w1 = w1 / scale1;

                            string Layer_name = _AGEN_mainform.Layer_odd;

                            for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot] != DBNull.Value)
                                {
                                    double xms = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x]);
                                    double yms = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y]);

                                    double Rot = Math.PI * Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot]) / 180;

                                    double x = xms - lr * (xps - xna) / scale1;
                                    double y = yms - lr * (yps - yna) / scale1;

                                    Circle C1 = new Circle(new Point3d(x, y, 0), Vector3d.ZAxis, 1);
                                    C1.TransformBy(Matrix3d.Rotation(Rot, Vector3d.ZAxis, new Point3d(xms, yms, 0)));


                                    Point3d Pt_ins = C1.Center;


                                    BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", _AGEN_mainform.NorthArrowMS, Pt_ins,
                                              1 / scale1, 0, Layer_name, new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());

                                    Functions.Populate_object_data_table_from_objectid(Tables1, block1.ObjectId, "Agen_Northarrow", Lista_val, Lista_type);

                                    if (Layer_name == _AGEN_mainform.Layer_odd)
                                    {
                                        Layer_name = _AGEN_mainform.Layer_even;
                                    }
                                    else
                                    {
                                        Layer_name = _AGEN_mainform.Layer_odd;
                                    }
                                }

                                else
                                {
                                    _AGEN_mainform.tpage_processing.Hide();
                                    MessageBox.Show("the sheet index file does not contain info about matchline rectangles... please check it");
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    return;
                                }
                            }


                            label_na.Visible = true;
                            Trans1.Commit();
                        }
                        else
                        {
                            MessageBox.Show(_AGEN_mainform.NorthArrowMS + " block not found in the drawing!");
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
            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();


        }

        private void button_insert_matchline_block_ms_Click(object sender, EventArgs e)
        {

            if (_AGEN_mainform.dt_sheet_index == null)
            {
                MessageBox.Show("no sheet index information loaded");
                return;
            }
            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("no sheet index information loaded");
                return;
            }

            if (_AGEN_mainform.dt_centerline == null)
            {
                MessageBox.Show("no centerline information loaded");
                return;
            }
            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("no centerline loaded");
                return;
            }

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

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }



            if (System.IO.Directory.Exists(ProjF) == true)
            {




                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }











            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
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
                _AGEN_mainform.tpage_processing.Show();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        if (BlockTable1.Has(_AGEN_mainform.matchline_block) == true)
                        {

                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            double scale1 = _AGEN_mainform.Vw_scale;

                            if (scale1 <= 0)
                            {
                                _AGEN_mainform.tpage_processing.Hide();
                                MessageBox.Show("please set up main viewport scale");
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
                                return;
                            }


                            #region object data stationing
                            Erase_matchline_blocks();
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                            Functions.Create_matchline_block_od_table();

                            List<object> Lista_val = new List<object>();
                            List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                            if (segment1 == "not defined") segment1 = "";
                            Lista_val.Add(segment1);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                            #endregion


                            Functions.Creaza_layer(_AGEN_mainform.Layer_even, 7, true);
                            Functions.Creaza_layer(_AGEN_mainform.Layer_odd, 7, true);


                            if (_AGEN_mainform.dt_station_equation != null)
                            {
                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                    _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);

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


                                            Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                            if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                            double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                            _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                        }
                                    }

                                    _AGEN_mainform.Poly3D.Erase();
                                }

                            }

                            string Layer_name = _AGEN_mainform.Layer_odd;

                            for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Height] != DBNull.Value
                                     && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Width] != DBNull.Value
                                      && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1] != DBNull.Value
                                       && _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2] != DBNull.Value)
                                {
                                    double xms = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_x]);
                                    double yms = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_y]);

                                    double m1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M1]);
                                    double m2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_M2]);

                                    double h1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Height]);
                                    double w1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_Width]);

                                    double Rot = Math.PI * Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_rot]) / 180;



                                    double x = xms - w1 / 2;
                                    double y = yms - h1 / 2;

                                    Circle C1 = new Circle(new Point3d(x, y, 0), Vector3d.ZAxis, 1);
                                    C1.TransformBy(Matrix3d.Rotation(Rot, Vector3d.ZAxis, new Point3d(xms, yms, 0)));


                                    Point3d Pt_ins = C1.Center;

                                    double dispm1 = Functions.Station_equation_ofV2(m1, _AGEN_mainform.dt_station_equation);
                                    double dispm2 = Functions.Station_equation_ofV2(m2, _AGEN_mainform.dt_station_equation);

                                    string StM1 = Functions.Get_chainage_from_double(dispm1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                    string StM2 = Functions.Get_chainage_from_double(dispm2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                    string mile1 = Functions.Get_String_Rounded(dispm1 / 5280, 1);
                                    string mile2 = Functions.Get_String_Rounded(dispm2 / 5280, 1);

                                    string Prev_file = "BEGIN STA.";
                                    string Next_file = "END STA.";

                                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 1)
                                    {
                                        if (i < _AGEN_mainform.dt_sheet_index.Rows.Count - 1)
                                        {
                                            if (_AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                            {
                                                Next_file = _AGEN_mainform.dt_sheet_index.Rows[i + 1][_AGEN_mainform.Col_dwg_name].ToString();
                                            }
                                        }

                                        if (i > 0)
                                        {
                                            if (_AGEN_mainform.dt_sheet_index.Rows[i - 1][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                            {
                                                Prev_file = _AGEN_mainform.dt_sheet_index.Rows[i - 1][_AGEN_mainform.Col_dwg_name].ToString();
                                            }
                                        }
                                    }

                                    System.Collections.Specialized.StringCollection Col_val = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Col_atr = new System.Collections.Specialized.StringCollection();

                                    Col_atr.Add("ATR_1");
                                    Col_val.Add(mile1);
                                    Col_atr.Add("ATR_2");
                                    Col_val.Add(Prev_file);
                                    Col_atr.Add("ATR_3");
                                    Col_val.Add(StM1);

                                    Col_atr.Add("ATR_4");
                                    Col_val.Add("");
                                    Col_atr.Add("ATR_5");
                                    Col_val.Add("");




                                    BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", _AGEN_mainform.matchline_block, Pt_ins,
                                               1 / scale1, Rot, Layer_name, Col_atr, Col_val);

                                    Functions.Stretch_block(block1, "Distance1", h1);

                                    x = xms + w1 / 2;
                                    y = yms - h1 / 2;

                                    Circle C2 = new Circle(new Point3d(x, y, 0), Vector3d.ZAxis, 1);
                                    C2.TransformBy(Matrix3d.Rotation(Rot, Vector3d.ZAxis, new Point3d(xms, yms, 0)));


                                    Pt_ins = C2.Center;

                                    Col_val = new System.Collections.Specialized.StringCollection();
                                    Col_atr = new System.Collections.Specialized.StringCollection();

                                    Col_atr.Add("ATR_1");
                                    Col_val.Add(mile2);
                                    Col_atr.Add("ATR_5");
                                    Col_val.Add(Next_file);
                                    Col_atr.Add("ATR_4");
                                    Col_val.Add(StM2);

                                    Col_atr.Add("ATR_2");
                                    Col_val.Add("");
                                    Col_atr.Add("ATR_3");
                                    Col_val.Add("");





                                    BlockReference block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", _AGEN_mainform.matchline_block, Pt_ins,
                                               1 / scale1, Rot, Layer_name, Col_atr, Col_val);

                                    Functions.Stretch_block(block2, "Distance1", h1);

                                    Functions.Populate_object_data_table_from_objectid(Tables1, block2.ObjectId, "Agen_mlblocks", Lista_val, Lista_type);

                                    if (Layer_name == _AGEN_mainform.Layer_odd)
                                    {
                                        Layer_name = _AGEN_mainform.Layer_even;
                                    }
                                    else
                                    {
                                        Layer_name = _AGEN_mainform.Layer_odd;
                                    }
                                }

                                else
                                {
                                    _AGEN_mainform.tpage_processing.Hide();
                                    MessageBox.Show("the sheet index file does not contain info about matchline rectangles... please check it");
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    set_enable_true();
                                    return;
                                }
                            }


                            label_mlb.Visible = true;
                            Trans1.Commit();
                        }
                        else
                        {
                            MessageBox.Show(_AGEN_mainform.matchline_block + " block not found in the drawing!");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();


        }

        private void zoom_to_object(ObjectId ObjId)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {
                            Entity Ent1 = Trans1.GetObject(ObjId, OpenMode.ForRead) as Entity;
                            if (Ent1 != null)
                            {

                                Point3d minx = Ent1.GeometricExtents.MinPoint;
                                Point3d maxx = Ent1.GeometricExtents.MaxPoint;

                                using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                                {

                                    int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                    //from here 2015 dlls:
                                    Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                    kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                    Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                    // to here 2015 dlls

                                    //from here 2013 dlls:

                                    //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                    // to here 2013 dlls

                                    if (view != null)
                                    {
                                        using (view)
                                        {

                                            view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);

                                            view.Zoom(0.95);//<--optional 
                                            GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                        }
                                    }
                                    Trans1.Commit();
                                }

                            }
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        private void zoom_to_Point(Point3d pt, double factor1)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - factor1, pt.Y - factor1, 0);
                            Point3d maxx = new Point3d(pt.X + factor1, pt.Y + factor1, 0);

                            using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                            {

                                int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                //from here 2015 dlls:
                                Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                // to here 2015 dlls

                                //from here 2013 dlls:

                                //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                // to here 2013 dlls

                                if (view != null)
                                {
                                    using (view)
                                    {

                                        view.ZoomExtents(minx, maxx);

                                        view.Zoom(0.95);//<--optional 
                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }
                                }
                                Trans1.Commit();
                            }


                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void delete_centerlines()
        {
            if (lista_del.Count > 0)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                try
                {

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            for (int i = 0; i < lista_del.Count; ++i)
                            {
                                try
                                {
                                    Entity Cl = Trans1.GetObject(lista_del[i], OpenMode.ForWrite) as Entity;
                                    if (Cl != null)
                                    {
                                        Cl.Erase();
                                    }
                                }
                                catch (System.Exception ex)
                                {

                                }


                            }

                            Trans1.Commit();
                        }

                        lista_del.Clear();
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");

            }
        }

        private void button_delete_3d_poly_Click(object sender, EventArgs e)
        {

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;


                delete_centerlines();
                Functions.create_backup(fisier_si);
                Populate_sheet_index_file(fisier_si);
                label_not_saved.Visible = false;
            }

        }


        private void button_delete_sheet_index_Click(object sender, EventArgs e)
        {

            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }

            if (MessageBox.Show("WARNING!!!\r\nThis will remove all sheet indexes from your drawing and remove all data from the Sheet Index Data Table.\r\nDo you want to continue?", "AGEN", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;

            }


            set_enable_false();
            try
            {

                _AGEN_mainform.dt_sheet_index = Functions.Creaza_sheet_index_datatable_structure();
                dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_sheet_index.EnableHeadersVisualStyles = false;
                _AGEN_mainform.tpage_processing.Show();
                if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                {
                    string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }
                    string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

                    Functions.create_backup(fisier_si);
                    Populate_sheet_index_file(fisier_si);

                    Erase_matchlines_templates();
                    Erase_viewports_templates();
                    Erase_Ms_blocks();
                    label_vp.Visible = false;
                    label_na.Visible = false;
                    label_mlb.Visible = false;

                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();
        }


        public void Hide_labels_at_load_project()
        {
            label_vp.Visible = false;
            label_na.Visible = false;
            label_mlb.Visible = false;
        }

        private void label_matchline_setup_Click(object sender, EventArgs e)
        {

            if (panel_dan.Visible == true)
            {
                panel_dan.Visible = false;
            }
            else
            {
                panel_dan.Visible = true;

            }

        }

        private void button_calc_rectangles_based_on_start_and_end_Click(object sender, EventArgs e)
        {


            string Col_M1 = "StaBeg";
            string Col_M2 = "StaEnd";

            string Col_X = "X";
            string Col_Y = "Y";
            string Col_rot = "Rotation";
            string Col_Width = "Width";
            string Col_Height = "Height";
            string Col_X1 = "X_Beg";
            string Col_Y1 = "Y_Beg";
            string Col_X2 = "X_End";
            string Col_Y2 = "Y_End";

            if (_AGEN_mainform.dt_sheet_index == null || _AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("no Sheet index data found. First load the data!", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _AGEN_mainform.tpage_processing.Hide();
                return;
            }


            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("no centerline data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _AGEN_mainform.tpage_processing.Hide();
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
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        set_enable_false();

                        Polyline3d poly3d = null;
                        Polyline poly2d = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                        }

                        System.Data.DataTable dt_cl = _AGEN_mainform.dt_centerline;

                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.COUNTRY == "USA")
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[i][Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[i][Col_M2] != DBNull.Value)
                                {
                                    double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_M1]);
                                    double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_M2]);

                                    double vertical_len = -1;

                                    if (_AGEN_mainform.dt_sheet_index.Rows[i][Col_X] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y] != DBNull.Value)
                                    {
                                        double xc1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_X]);
                                        double yc1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_Y]);
                                        Point3d pt_on_poly1 = poly2d.GetClosestPointTo(new Point3d(xc1, yc1, poly2d.Elevation), Vector3d.ZAxis, false);
                                        Polyline poly_temp1 = new Polyline();
                                        poly_temp1.AddVertexAt(0, new Point2d(pt_on_poly1.X, pt_on_poly1.Y), 0, 0, 0);
                                        poly_temp1.AddVertexAt(1, new Point2d(xc1, yc1), 0, 0, 0);
                                        vertical_len = poly_temp1.Length;
                                    }

                                    if (M2 <= M1)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + "is bigger or equal than " + M2.ToString(), "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }

                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        if (M1 > poly3d.Length || M2 > poly3d.Length)
                                        {
                                            set_enable_true();
                                            MessageBox.Show("Matchline " + M1.ToString() + " or " + M2.ToString() + "\r\nis larger than centerline\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            _AGEN_mainform.tpage_processing.Hide();
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        if (M1 > poly2d.Length || M2 > poly2d.Length)
                                        {
                                            set_enable_true();
                                            MessageBox.Show("Matchline " + M1.ToString() + " or " + M2.ToString() + "\r\nis larger than centerline\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            _AGEN_mainform.tpage_processing.Hide();
                                            return;
                                        }
                                    }


                                    if (M1 < 0 || M2 < 0)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + " or " + M2.ToString() + "\r\nis smaller than 0\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }

                                    if (M1 == M2)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + " and " + M2.ToString() + "\r\nare identical!\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }

                                    Point3d point1 = new Point3d();
                                    Point3d point2 = new Point3d();

                                    if (_AGEN_mainform.Project_type == "3D")
                                    {
                                        point1 = poly3d.GetPointAtDist(M1);
                                        point2 = poly3d.GetPointAtDist(M2);
                                    }
                                    else
                                    {
                                        point1 = poly2d.GetPointAtDist(M1);
                                        point2 = poly2d.GetPointAtDist(M2);
                                    }


                                    Polyline rect1 = create_rectangle_Matchline(point1, point2, 7);

                                    double xc2 = (rect1.GetPoint3dAt(0).X + rect1.GetPoint3dAt(2).X) / 2;
                                    double yc2 = (rect1.GetPoint3dAt(0).Y + rect1.GetPoint3dAt(2).Y) / 2;

                                    if (vertical_len > 0)
                                    {
                                        Point3d pt_on_poly2 = poly2d.GetClosestPointTo(new Point3d(xc2, yc2, poly2d.Elevation), Vector3d.ZAxis, false);                                      
                                        using (Polyline poly_temp2 = new Polyline())
                                        {
                                            poly_temp2.AddVertexAt(0, new Point2d(pt_on_poly2.X, pt_on_poly2.Y), 0, 0, 0);
                                            poly_temp2.AddVertexAt(1, new Point2d(xc2, yc2), 0, 0, 0);
                                            poly_temp2.TransformBy(Matrix3d.Scaling(vertical_len / poly_temp2.Length, poly_temp2.StartPoint));
                                            xc2 = poly_temp2.EndPoint.X;
                                            yc2 = poly_temp2.EndPoint.Y;
                                        }
                                    }

                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X] = xc2;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y] = yc2;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_rot] = Functions.GET_Bearing_rad(rect1.GetPoint3dAt(1).X, rect1.GetPoint3dAt(1).Y, rect1.GetPoint3dAt(2).X, rect1.GetPoint3dAt(2).Y) * 180 / Math.PI;

                                    Point3d p0 = rect1.GetPointAtParameter(0);
                                    Point3d p1 = rect1.GetPointAtParameter(1);
                                    Point3d p2 = rect1.GetPointAtParameter(2);

                                    double d12 = Math.Pow(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2), 0.5);
                                    double d01 = Math.Pow(Math.Pow(p1.X - p0.X, 2) + Math.Pow(p1.Y - p0.Y, 2), 0.5);

                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Width] = d12;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Height] = d01;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X1] = point1.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y1] = point1.Y;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X2] = point2.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y2] = point2.Y;

                                    label_not_saved.Visible = true;
                                }
                                else
                                {
                                    set_enable_true();
                                    MessageBox.Show("Matchline value is not specified correctly.\r\nReview your sheet index file\r\nbegin and end station columns", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }

                            if (_AGEN_mainform.COUNTRY == "CANADA")
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows[i][Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[i][Col_M2] != DBNull.Value)
                                {
                                    double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_M1]);
                                    double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i][Col_M2]);

                                    if (M2 <= M1)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + "is bigger or equal than " + M2.ToString(), "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }


                                    if (M1 == M2)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + " and " + M2.ToString() + "\r\nare identical!\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }

                                    double m1_2d = -1;
                                    double m2_2d = -1;

                                    for (int j = 0; j < dt_cl.Rows.Count - 1; ++j)
                                    {
                                        if (dt_cl.Rows[j]["3DSta"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                           dt_cl.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                           dt_cl.Rows[j]["X"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["X"])) == true &&
                                           dt_cl.Rows[j]["Y"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j]["Y"])) == true &&
                                            dt_cl.Rows[j]["Z"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true &&
                                            dt_cl.Rows[j]["X"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["X"])) == true &&
                                           dt_cl.Rows[j + 1]["Y"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Y"])) == true &&
                                            dt_cl.Rows[j + 1]["Z"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(dt_cl.Rows[j + 1]["Z"])) == true)
                                        {
                                            double sta1 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j]["3DSta"]).Replace("+", ""));
                                            double sta2 = Convert.ToDouble(Convert.ToString(dt_cl.Rows[j + 1]["3DSta"]).Replace("+", ""));


                                            double x1 = Convert.ToDouble(dt_cl.Rows[j]["X"]);
                                            double y1 = Convert.ToDouble(dt_cl.Rows[j]["Y"]);
                                            double z1 = Convert.ToDouble(dt_cl.Rows[j]["Z"]);
                                            double x2 = Convert.ToDouble(dt_cl.Rows[j + 1]["X"]);
                                            double y2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Y"]);
                                            double z2 = Convert.ToDouble(dt_cl.Rows[j + 1]["Z"]);


                                            if (M1 >= sta1 && M1 <= sta2)
                                            {

                                                double x = x1 + (x2 - x1) * (M1 - sta1) / (sta2 - sta1);
                                                double y = y1 + (y2 - y1) * (M1 - sta1) / (sta2 - sta1);
                                                double z = z1 + (z2 - z1) * (M1 - sta1) / (sta2 - sta1);

                                                Point3d pt1 = new Point3d(x, y, 0);

                                                m1_2d = poly2d.GetDistAtPoint(poly2d.GetClosestPointTo(pt1, Vector3d.ZAxis, false));

                                            }

                                            if (M2 >= sta1 && M2 <= sta2)
                                            {

                                                double x = x1 + (x2 - x1) * (M2 - sta1) / (sta2 - sta1);
                                                double y = y1 + (y2 - y1) * (M2 - sta1) / (sta2 - sta1);
                                                double z = z1 + (z2 - z1) * (M2 - sta1) / (sta2 - sta1);
                                                Point3d pt2 = new Point3d(x, y, 0);
                                                m2_2d = poly2d.GetDistAtPoint(poly2d.GetClosestPointTo(pt2, Vector3d.ZAxis, false));


                                            }

                                        }

                                    }


                                    if (m1_2d > poly2d.Length || m2_2d > poly2d.Length)
                                    {
                                        set_enable_true();
                                        MessageBox.Show("Matchline " + M1.ToString() + " or " + M2.ToString() + "\r\nis larger than centerline\r\nReview your data", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        _AGEN_mainform.tpage_processing.Hide();
                                        return;
                                    }


                                    Point3d point1 = poly2d.GetPointAtDist(m1_2d);
                                    Point3d point2 = poly2d.GetPointAtDist(m2_2d);

                                    Polyline rect1 = create_rectangle_Matchline(point1, point2, 7);


                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X] = (rect1.GetPoint3dAt(0).X + rect1.GetPoint3dAt(2).X) / 2;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y] = (rect1.GetPoint3dAt(0).Y + rect1.GetPoint3dAt(2).Y) / 2;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_rot] = Functions.GET_Bearing_rad(rect1.GetPoint3dAt(1).X, rect1.GetPoint3dAt(1).Y, rect1.GetPoint3dAt(2).X, rect1.GetPoint3dAt(2).Y) * 180 / Math.PI;

                                    Point3d p0 = rect1.GetPointAtParameter(0);
                                    Point3d p1 = rect1.GetPointAtParameter(1);
                                    Point3d p2 = rect1.GetPointAtParameter(2);


                                    double d12 = Math.Pow(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2), 0.5);
                                    double d01 = Math.Pow(Math.Pow(p1.X - p0.X, 2) + Math.Pow(p1.Y - p0.Y, 2), 0.5);

                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Width] = d12;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Height] = d01;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X1] = point1.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y1] = point1.Y;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_X2] = point2.X;
                                    _AGEN_mainform.dt_sheet_index.Rows[i][Col_Y2] = point2.Y;

                                    label_not_saved.Visible = true;

                                }
                                else
                                {
                                    set_enable_true();
                                    MessageBox.Show("Matchline value is not specified correctly.\r\nReview your sheet index file\r\nbegin and end station columns", "Sheet index adjustment", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }



                        }


                        if (poly3d != null)
                        {
                            poly3d.Erase();
                            Trans1.Commit();
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



        private void button_open_sheet_index_xl_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {

                string file_de_procesat = "";


                if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                {
                    string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }

                    string fisier_sheet_index = ProjF + _AGEN_mainform.sheet_index_excel_name;

                    if (System.IO.File.Exists(fisier_sheet_index) == true)

                    {
                        file_de_procesat = fisier_sheet_index;

                    }
                }

                if (System.IO.File.Exists(file_de_procesat) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the block sheet inderx data file does not exist");
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
                Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(file_de_procesat);




            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();


        }

        public void add_segment_name(string segmentname)
        {
            label_sheet_index_data.Text = "Sheet Index Data" + " " + segmentname;
        }

        private void button_load_segment_sheet_index_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_setup.Build_sheet_index_dt_from_excel();

        }

        private void button_scan_Click(object sender, EventArgs e)
        {

            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("no project loaded");
                return;
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("the centerline data file does not exist");
                return;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);
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
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect rectangles viewport creation:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status == PromptStatus.OK)
                        {

                            #region read existing sheet index
                            string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;
                            if (System.IO.File.Exists(fisier_si) == true)
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
                                Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                                Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                                try
                                {
                                    if (Excel1 == null)
                                    {
                                        MessageBox.Show("PROBLEM WITH EXCEL!");
                                        return;
                                    }

                                    Workbook2 = Excel1.Workbooks.Open(fisier_si);
                                    W2 = Workbook2.Worksheets[1];
                                    _AGEN_mainform.dt_sheet_index = Functions.Build_Data_table_sheet_index_from_excel(W2, _AGEN_mainform.Start_row_Sheet_index + 1);
                                    Workbook2.Close();
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
                                    MessageBox.Show(ex.Message);
                                }
                                finally
                                {

                                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                                    if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                                }
                            }
                            else
                            {
                                _AGEN_mainform.dt_sheet_index = Functions.Creaza_sheet_index_datatable_structure();
                            }
                            #endregion


                            #region  LOAD EXISTING CL
                            _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                            _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            #endregion  

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline rect1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                if (rect1 != null)
                                {
                                    if (rect1.NumberOfVertices > 5 || rect1.NumberOfVertices < 4)
                                    {
                                        MessageBox.Show("The rectangle can only have 4 or 5 vertices\r\nFix your rectangle and try again");
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        set_enable_true();

                                        return;
                                    }

                                    Point3d pt1 = new Point3d(0, 1, 2);
                                    Point3d pt2 = new Point3d(0, 1, 2);
                                    Point3d pt3 = new Point3d(0, 1, 2);
                                    Point3d pt4 = new Point3d(0, 1, 2);

                                    if (rect1.NumberOfVertices == 4)
                                    {
                                        pt1 = rect1.GetPointAtParameter(0);
                                        pt2 = rect1.GetPointAtParameter(1);
                                        pt3 = rect1.GetPointAtParameter(2);
                                        pt4 = rect1.GetPointAtParameter(3);
                                    }

                                    if (rect1.NumberOfVertices == 5)
                                    {
                                        pt1 = rect1.GetPointAtParameter(0);
                                        pt2 = rect1.GetPointAtParameter(1);
                                        pt3 = rect1.GetPointAtParameter(2);
                                        pt4 = rect1.GetPointAtParameter(3);

                                        Point3d pt5 = rect1.GetPointAtParameter(4);

                                        double dist1 = get_distance(pt1, pt2);
                                        double dist2 = get_distance(pt1, pt3);
                                        double dist3 = get_distance(pt1, pt4);
                                        double dist4 = get_distance(pt2, pt3);
                                        double dist5 = get_distance(pt2, pt4);
                                        double dist6 = get_distance(pt3, pt4);

                                        if (Math.Round(dist1, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist2, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist3, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist4, 0) == 0)
                                        {
                                            pt2 = pt5;
                                        }
                                        if (Math.Round(dist5, 0) == 0)
                                        {
                                            pt2 = pt5;
                                        }
                                        if (Math.Round(dist6, 0) == 0)
                                        {
                                            pt3 = pt5;
                                        }
                                    }

                                    if (pt1 != new Point3d(0, 1, 2) && pt2 != new Point3d(0, 1, 2) && pt3 != new Point3d(0, 1, 2) && pt4 != new Point3d(0, 1, 2))
                                    {
                                        double dist1 = get_distance(pt1, pt2);
                                        double dist2 = get_distance(pt2, pt3);
                                        double dist3 = get_distance(pt3, pt4);
                                        double dist4 = get_distance(pt4, pt1);
                                        if (Math.Abs(dist1 - dist3) < 5 && Math.Abs(dist2 - dist4) < 5)
                                        {
                                            // i assume 2 to 3 is the length 1 to 3 is the width
                                            double bear1 = Functions.GET_Bearing_rad(pt2.X, pt2.Y, pt3.X, pt3.Y);
                                            double xc = (pt1.X + pt3.X) / 2;
                                            double yc = (pt1.Y + pt3.Y) / 2;
                                            double l1 = dist2;
                                            double h1 = dist1;


                                            Polyline poly_start = new Polyline();
                                            poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                            poly_start.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                            poly_start.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                            Polyline poly_end = new Polyline();
                                            poly_end.AddVertexAt(0, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                            poly_end.AddVertexAt(1, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                            poly_end.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                            Polyline poly_top = new Polyline();
                                            poly_top.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                            poly_top.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                            poly_top.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                            Polyline poly_bottom = new Polyline();
                                            poly_bottom.AddVertexAt(0, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                            poly_bottom.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                            poly_bottom.Elevation = _AGEN_mainform.Poly2D.Elevation;




                                            if (dist1 > dist2)
                                            {
                                                // i assume 1 to 2 is the length 2 to 3 is the width
                                                bear1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                                xc = (pt1.X + pt3.X) / 2;
                                                yc = (pt1.Y + pt3.Y) / 2;
                                                l1 = dist1;
                                                h1 = dist2;

                                                poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                                poly_start.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                                poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                                poly_top = new Polyline();
                                                poly_top.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_top.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                                poly_top.Elevation = _AGEN_mainform.Poly2D.Elevation;

                                                poly_bottom = new Polyline();
                                                poly_bottom.AddVertexAt(0, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                                poly_bottom.AddVertexAt(1, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                                poly_bottom.Elevation = _AGEN_mainform.Poly2D.Elevation;



                                            }

                                            string dwg_name = "";

                                            #region object data
                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), rect1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                            {
                                                if (Records1 != null)
                                                {
                                                    if (Records1.Count > 0)
                                                    {

                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                            for (int j = 0; j < Record1.Count; ++j)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                                string Nume_field = Field_def1.Name;
                                                                object valoare1 = Record1[j].StrValue;
                                                                if (Nume_field.ToLower() == "drawingnum")
                                                                {
                                                                    dwg_name = Convert.ToString(valoare1);
                                                                    j = Record1.Count;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            int index_si = _AGEN_mainform.dt_sheet_index.Rows.Count;

                                            for (int k = 0; k < _AGEN_mainform.dt_sheet_index.Rows.Count; ++k)
                                            {
                                                if (_AGEN_mainform.dt_sheet_index.Rows[k]["DwgNo"] != DBNull.Value)
                                                {
                                                    string dwg2 = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[k]["DwgNo"]);

                                                    if (dwg2.ToLower() == dwg_name.ToLower())
                                                    {
                                                        index_si = k;
                                                        k = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                    }
                                                }
                                            }



                                            if (index_si == _AGEN_mainform.dt_sheet_index.Rows.Count)
                                            {
                                                _AGEN_mainform.dt_sheet_index.Rows.Add();
                                            }
                                            else
                                            {
                                                if (MessageBox.Show(dwg_name + " \r\nhas been found in sheet index.\r\nDo you want to replace it?", "agen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                                {
                                                    _AGEN_mainform.dt_sheet_index.Rows.Add();
                                                    index_si = _AGEN_mainform.dt_sheet_index.Rows.Count - 1;
                                                }
                                            }

                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Rotation"] = bear1 * 180 / Math.PI;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["X"] = xc;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Y"] = yc;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["DwgNo"] = dwg_name;


                                            Point3dCollection col_start = Functions.Intersect_on_both_operands(_AGEN_mainform.Poly2D, poly_start);
                                            Point3dCollection col_end = Functions.Intersect_on_both_operands(_AGEN_mainform.Poly2D, poly_end);

                                            Point3dCollection col_top = Functions.Intersect_on_both_operands(_AGEN_mainform.Poly2D, poly_top);
                                            Point3dCollection col_bottom = Functions.Intersect_on_both_operands(_AGEN_mainform.Poly2D, poly_bottom);

                                            double sta_beg = -1;
                                            double sta_end = -1;
                                            Point3d p1 = new Point3d();
                                            Point3d p2 = new Point3d();


                                            if (col_start.Count > 0)
                                            {
                                                Point3d point1 = col_start[0];
                                                Point3d p10 = _AGEN_mainform.Poly2D.GetClosestPointTo(point1, Vector3d.ZAxis, false);

                                                double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(p10);
                                                if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    sta_beg = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);


                                                }
                                                else
                                                {
                                                    double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                                    double b1 = -1.23456;
                                                    sta_beg = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, p10, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                                }

                                                p1 = p10;

                                            }


                                            if (col_end.Count > 0)
                                            {
                                                Point3d point1 = col_end[0];
                                                Point3d p10 = _AGEN_mainform.Poly2D.GetClosestPointTo(point1, Vector3d.ZAxis, false);

                                                double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(p10);
                                                if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    sta_end = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);


                                                }
                                                else
                                                {
                                                    double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                                    double b2 = -1.23456;
                                                    sta_end = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, p10, dist_2d1, _AGEN_mainform.dt_centerline, ref b2);
                                                }

                                                p2 = p10;

                                            }


                                            if (col_start.Count == 0 && col_end.Count == 0)
                                            {
                                                if (col_top.Count > 0)
                                                {

                                                    Point3d point1 = col_top[0];
                                                    Point3d p10 = _AGEN_mainform.Poly2D.GetClosestPointTo(point1, Vector3d.ZAxis, false);

                                                    double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(p10);
                                                    if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                                    if (_AGEN_mainform.COUNTRY == "USA")
                                                    {
                                                        sta_beg = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);


                                                    }
                                                    else
                                                    {
                                                        double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                                        double b1 = -1.23456;
                                                        sta_beg = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, p10, dist_2d1, _AGEN_mainform.dt_centerline, ref b1);
                                                    }

                                                    p1 = p10;

                                                }


                                                if (col_bottom.Count > 0)
                                                {
                                                    Point3d point1 = col_bottom[0];
                                                    Point3d p10 = _AGEN_mainform.Poly2D.GetClosestPointTo(point1, Vector3d.ZAxis, false);

                                                    double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(p10);
                                                    if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                                    if (_AGEN_mainform.COUNTRY == "USA")
                                                    {
                                                        sta_end = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);


                                                    }
                                                    else
                                                    {
                                                        double dist_2d1 = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                                        double b2 = -1.23456;
                                                        sta_end = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, p10, dist_2d1, _AGEN_mainform.dt_centerline, ref b2);
                                                    }

                                                    p2 = p10;

                                                }

                                                if (col_top.Count > 0 && col_bottom.Count > 0)
                                                {
                                                    double T = h1;
                                                    h1 = l1;
                                                    l1 = T;
                                                }


                                            }

                                            if (sta_beg != -1 && sta_end != -1)
                                            {
                                                if (sta_beg > sta_end)
                                                {
                                                    double t = sta_beg;
                                                    sta_beg = sta_end;
                                                    sta_end = t;
                                                    Point3d tt1 = p1;
                                                    p1 = p2;
                                                    p2 = tt1;
                                                    double r1 = bear1 * 180 / Math.PI;

                                                    if (r1 - 180 < 0)
                                                    {
                                                        r1 = r1 + 180;
                                                    }
                                                    else
                                                    {
                                                        r1 = r1 - 180;
                                                    }

                                                    _AGEN_mainform.dt_sheet_index.Rows[index_si]["Rotation"] = r1;
                                                }
                                            }
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Height"] = h1;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Width"] = l1;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["StaBeg"] = sta_beg;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["StaEnd"] = sta_end;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["X_Beg"] = p1.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Y_Beg"] = p1.Y;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["X_End"] = p2.X;
                                            _AGEN_mainform.dt_sheet_index.Rows[index_si]["Y_End"] = p2.Y;
                                        }

                                    }

                                }
                            }
                            _AGEN_mainform.Poly3D.Erase();
                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {
                                Functions.create_backup(fisier_si);
                                Populate_sheet_index_file(fisier_si);

                                dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;

                            }

                            Trans1.Commit();

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

        private void button_draw_manual_Click(object sender, EventArgs e)

        {

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


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }



            if (_AGEN_mainform.Vw_height == 0 || _AGEN_mainform.Vw_width == 0)
            {
                MessageBox.Show("you do not have the dimensions for the matchline rectangles\r\nOperation aborted");


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


                return;
            }

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you do not have picked the centerline\r\noperation aborted");


                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();


                return;
            }


            double poly_length = 0;

            Ag.WindowState = FormWindowState.Minimized;


            set_enable_false();

            Erase_viewports_templates();

            Create_ML_object_data();




            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_AGEN_mainform.Layer_name_ML_rectangle, 4, false);

                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                    poly_length = _AGEN_mainform.Poly3D.Length;

                    delete_centerlines();

                    zoom_to_object(_AGEN_mainform.Poly3D.ObjectId);
                    lista_del.Add(_AGEN_mainform.Poly3D.ObjectId);
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {



                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                        _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);



                        #region station_eq
                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
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


                                    Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                    double p1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                    if (p1 > _AGEN_mainform.Poly3D.EndParam) p1 = _AGEN_mainform.Poly3D.EndParam;

                                    double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(p1);
                                    _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                }
                            }
                        }
                        #endregion





                        if (_AGEN_mainform.dt_sheet_index == null)
                        {
                            _AGEN_mainform.dt_sheet_index = Functions.Creaza_sheet_index_datatable_structure();
                        }

                        string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();
                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                        }
                        else
                        {
                            if (Scale1.Contains(":") == true)
                            {
                                Scale1 = Scale1.Replace("1:", "");
                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                                }
                            }
                            else
                            {
                                string inch = "\u0022";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                                }

                                inch = "\u0094";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                                }

                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                                }
                            }
                        }






                        int Colorindex = 1;




                    l123:

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1m = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease pick start location:");


                        PP1m.AllowNone = false;
                        Result_point_m1 = Editor1.GetPoint(PP1m);

                        if (Result_point_m1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                        {
                            Editor1.WriteMessage("\nCommand:");
                            Trans1.Commit();
                            goto end1;
                        }
                        Point3d Point1 = Result_point_m1.Value;




                        Alignment_mdi.Jig_rectangle_viewport_manual Jig2 = new Alignment_mdi.Jig_rectangle_viewport_manual();
                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m2 = Jig2.StartJig(_AGEN_mainform.Vw_scale, _AGEN_mainform.Vw_height, Point1, 10);

                        if (Result_point_m2.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                        {
                            Editor1.WriteMessage("\nCommand:");
                            Trans1.Commit();
                            goto end1;
                        }

                        Point3d Point2 = Result_point_m2.Value;

                        Polyline Rectangle2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                        Rectangle2 = create_rectangle_Matchline(Point1, Point2, Colorindex);
                        Rectangle2.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                        BTrecord.AppendEntity(Rectangle2);
                        Trans1.AddNewlyCreatedDBObject(Rectangle2, true);



                        Trans1.TransactionManager.QueueForGraphicsFlush();

                        Polyline Poly_left = new Polyline();

                        Poly_left.AddVertexAt(0, Rectangle2.GetPoint2dAt(0), 0, 0, 0);
                        Poly_left.AddVertexAt(1, Rectangle2.GetPoint2dAt(1), 0, 0, 0);
                        Poly_left.Elevation = _AGEN_mainform.Poly2D.Elevation;

                        Point3dCollection col_left = Functions.Intersect_with_extend(Poly_left, _AGEN_mainform.Poly2D);
                        double dist1 = -1;
                        if (col_left.Count > 0)
                        {
                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(col_left[0], Vector3d.ZAxis, false));
                            if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                            dist1 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                            Point1 = _AGEN_mainform.Poly3D.GetPointAtParameter(param1);

                        }

                        Polyline Poly_right = new Polyline();

                        Poly_right.AddVertexAt(0, Rectangle2.GetPoint2dAt(3), 0, 0, 0);
                        Poly_right.AddVertexAt(1, Rectangle2.GetPoint2dAt(2), 0, 0, 0);
                        Poly_right.Elevation = _AGEN_mainform.Poly2D.Elevation;


                        Point3dCollection col_right = Functions.Intersect_with_extend(Poly_right, _AGEN_mainform.Poly2D);
                        double dist2 = -1;
                        if (col_right.Count > 0)
                        {
                            double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(_AGEN_mainform.Poly2D.GetClosestPointTo(col_right[0], Vector3d.ZAxis, false));
                            if (param2 > _AGEN_mainform.Poly3D.EndParam) param2 = _AGEN_mainform.Poly3D.EndParam;
                            dist2 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param2);
                            Point2 = _AGEN_mainform.Poly3D.GetPointAtParameter(param2);

                        }


                        _AGEN_mainform.dt_sheet_index.Rows.Add();
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = dist1;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = dist2;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle2.ObjectId.Handle.Value.ToString();
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle2.GetPoint3dAt(0).X + Rectangle2.GetPoint3dAt(2).X) / 2;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle2.GetPoint3dAt(0).Y + Rectangle2.GetPoint3dAt(2).Y) / 2;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle2.GetPoint3dAt(1).X, Rectangle2.GetPoint3dAt(1).Y, Rectangle2.GetPoint3dAt(2).X, Rectangle2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle2.GetPoint3dAt(1).DistanceTo(Rectangle2.GetPoint3dAt(2));
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle2.GetPoint3dAt(0).DistanceTo(Rectangle2.GetPoint3dAt(1));
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1.X;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1.Y;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2.X;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2.Y;


                        Colorindex = Colorindex + 1;
                        if (Colorindex > 7) Colorindex = 1;



                        goto l123;

                    }


                end1:
                    if (_AGEN_mainform.dt_sheet_index != null)
                    {
                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {
                            Populate_data_table_matchline_file_names();




                            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;


                                round_sheet_index_data_table(poly_length);

                                Append_ML_object_data();
                                dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;


                                delete_centerlines();
                                label_not_saved.Visible = true;
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

            Ag.WindowState = FormWindowState.Normal;
        }


        private double get_distance(Point3d pt1, Point3d pt2)
        {
            double x1 = pt1.X;
            double y1 = pt1.Y;
            double x2 = pt2.X;
            double y2 = pt2.Y;

            return Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);

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

        private void comboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.current_segment = comboBox_segment_name.Text;
            _AGEN_mainform.tpage_setup.set_combobox_segment_name();


        }

        private void button_pick_middle_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_length.Text) == false)
            {
                return;
            }
            double rect_len = Math.Abs(Convert.ToDouble(textBox_length.Text));

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


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }



            if (_AGEN_mainform.Vw_height == 0 || _AGEN_mainform.Vw_width == 0)
            {
                MessageBox.Show("you do not have the dimensions for the matchline rectangles\r\nOperation aborted");


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


                return;
            }

            if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you do not have picked the centerline\r\noperation aborted");


                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();

                _AGEN_mainform.tpage_viewport_settings.Hide();
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


                _AGEN_mainform.tpage_setup.Show();


                return;
            }


            double poly_length = 0;

            Ag.WindowState = FormWindowState.Minimized;


            set_enable_false();

            Erase_viewports_templates();

            Create_ML_object_data();



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



                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                        _AGEN_mainform.Poly2D = Functions.Build_2d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                        _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                        poly_length = _AGEN_mainform.Poly3D.Length;
                        lista_del.Add(_AGEN_mainform.Poly3D.ObjectId);


                        #region station_eq
                        if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
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


                                    Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                    double p1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                    if (p1 > _AGEN_mainform.Poly3D.EndParam) p1 = _AGEN_mainform.Poly3D.EndParam;

                                    double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(p1);
                                    _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                }
                            }
                        }
                        #endregion

                        if (_AGEN_mainform.dt_sheet_index == null)
                        {
                            _AGEN_mainform.dt_sheet_index = Functions.Creaza_sheet_index_datatable_structure();
                        }


                        string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();

                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                        }
                        else
                        {
                            if (Scale1.Contains(":") == true)
                            {
                                Scale1 = Scale1.Replace("1:", "");
                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                                }
                            }
                            else
                            {
                                string inch = "\u0022";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                                }

                                inch = "\u0094";

                                if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                                {
                                    Scale1 = Scale1.Replace("1" + inch + "=", "");
                                    Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                                }

                                if (Functions.IsNumeric(Scale1) == true)
                                {
                                    _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                                }
                            }
                        }






                        int Colorindex = 1;




                    l123:

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point_m1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1m = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\n pick the middle location:");


                        PP1m.AllowNone = false;
                        Result_point_m1 = Editor1.GetPoint(PP1m);

                        if (Result_point_m1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                        {
                            Editor1.WriteMessage("\nCommand:");
                            Trans1.Commit();

                            goto end1;
                        }

                        Point3d Pointm = Result_point_m1.Value;

                        Point3d ptm_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(Pointm, Vector3d.ZAxis, false);
                        double paramm = _AGEN_mainform.Poly2D.GetParameterAtPoint(ptm_on_2d);
                        double stam = _AGEN_mainform.Poly3D.GetDistanceAtParameter(paramm);

                        double m1 = stam - 0.5 * rect_len;
                        double m2 = stam + 0.5 * rect_len;

                        if (m1 < 0 && m2 > _AGEN_mainform.Poly3D.Length)
                        {
                            MessageBox.Show("the length you specified is not fitting with the middle point\r\nmiddle-0.5*length < 0 or middle+0.5*length > centerline length\r\noperation aborted");
                            _AGEN_mainform.tpage_processing.Hide();
                            _AGEN_mainform.tpage_blank.Hide();
                            _AGEN_mainform.tpage_viewport_settings.Hide();
                            _AGEN_mainform.tpage_tblk_attrib.Hide();
                            _AGEN_mainform.tpage_setup.Hide();
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

                            _AGEN_mainform.tpage_sheetindex.Show();
                            Ag.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            return;
                        }



                        Point3d Point1 = _AGEN_mainform.Poly3D.GetPointAtDist(m1);
                        Point3d Point2 = _AGEN_mainform.Poly3D.GetPointAtDist(m2);



                        Polyline Rectangle2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                        Rectangle2 = create_rectangle_Matchline(Point1, Point2, Colorindex);
                        Rectangle2.Layer = _AGEN_mainform.Layer_name_ML_rectangle;

                        BTrecord.AppendEntity(Rectangle2);
                        Trans1.AddNewlyCreatedDBObject(Rectangle2, true);

                        Trans1.TransactionManager.QueueForGraphicsFlush();


                        double m1s = Functions.Station_equation_ofV2(m1, _AGEN_mainform.dt_station_equation);
                        double m2s = Functions.Station_equation_ofV2(m2, _AGEN_mainform.dt_station_equation);


                        _AGEN_mainform.dt_sheet_index.Rows.Add();
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M1] = m1s;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_M2] = m2s;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_handle] = Rectangle2.ObjectId.Handle.Value.ToString();
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_x] = (Rectangle2.GetPoint3dAt(0).X + Rectangle2.GetPoint3dAt(2).X) / 2;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_y] = (Rectangle2.GetPoint3dAt(0).Y + Rectangle2.GetPoint3dAt(2).Y) / 2;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle2.GetPoint3dAt(1).X, Rectangle2.GetPoint3dAt(1).Y, Rectangle2.GetPoint3dAt(2).X, Rectangle2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Width] = Rectangle2.GetPoint3dAt(1).DistanceTo(Rectangle2.GetPoint3dAt(2));
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1][_AGEN_mainform.Col_Height] = Rectangle2.GetPoint3dAt(0).DistanceTo(Rectangle2.GetPoint3dAt(1));
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_Beg"] = Point1.X;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_Beg"] = Point1.Y;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["X_End"] = Point2.X;
                        _AGEN_mainform.dt_sheet_index.Rows[_AGEN_mainform.dt_sheet_index.Rows.Count - 1]["Y_End"] = Point2.Y;

                        Colorindex = Colorindex + 1;
                        if (Colorindex > 7) Colorindex = 1;

                        goto l123;

                    }


                end1:
                    if (_AGEN_mainform.dt_sheet_index != null)
                    {
                        if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {
                            Populate_data_table_matchline_file_names();




                            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;


                                round_sheet_index_data_table(poly_length);

                                Append_ML_object_data();
                                dataGridView_sheet_index.DataSource = _AGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;


                                delete_centerlines();
                                label_not_saved.Visible = true;
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

            Ag.WindowState = FormWindowState.Normal;

            set_enable_true();


        }


    }



}
