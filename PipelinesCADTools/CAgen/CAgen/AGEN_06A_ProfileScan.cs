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

namespace Alignment_mdi
{

    public partial class AGEN_ProfileScan : Form
    {
        System.Data.DataTable dt_hl = null;
        string Col_BackSta = "BackSta";
        string Col_AheadSta = "AheadSta";

        int rec_no = 0;
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_cl_intersector_scan);
            lista_butoane.Add(button_read_prof_to_xl);
            lista_butoane.Add(button_show_scan_and_draw_profile);

            lista_butoane.Add(label_scan_profile_data);
            lista_butoane.Add(button_generate_profile3D_xls);
            lista_butoane.Add(button_calc_low_high);
            lista_butoane.Add(button_draw_Mleader);
            lista_butoane.Add(button_export_to_excel_high_low);

            foreach (Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_cl_intersector_scan);
            lista_butoane.Add(button_read_prof_to_xl);
            lista_butoane.Add(button_show_scan_and_draw_profile);

            lista_butoane.Add(label_scan_profile_data);
            lista_butoane.Add(button_generate_profile3D_xls);
            lista_butoane.Add(button_calc_low_high);
            lista_butoane.Add(button_draw_Mleader);
            lista_butoane.Add(button_export_to_excel_high_low);
            foreach (Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        public AGEN_ProfileScan()
        {
            InitializeComponent();
        }

        private void button_scan_for_prof_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.prof_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.prof_excel_name + " file");
                return;
            }

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            ObjectId[] Empty_array = null;
            Editor1.SetImpliedSelection(Empty_array);
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            ObjectId ObjID_FAIL = ObjectId.Null;





            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }

            if (System.IO.Directory.Exists(ProjF) == false)
            {
                MessageBox.Show("no project database folder found\r\nOperation aborted");
                return;
            }


            string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                set_enable_true();
                MessageBox.Show("the centerline data file does not exist");
                return;
            }


            if (_AGEN_mainform.dt_centerline == null)
            {
                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
            }
            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("the centerline data!");
                return;
            }




            double depth1 = 0;
            double diam1 = 0;

            if (Functions.IsNumeric(textBox_cover.Text) == true)
            {
                depth1 = Convert.ToDouble(textBox_cover.Text);
            }

            if (Functions.IsNumeric(textBox_pipe_diam.Text) == true)
            {
                diam1 = Convert.ToDouble(textBox_pipe_diam.Text);
            }

            if (_AGEN_mainform.COUNTRY == "USA")
            {
                diam1 = diam1 / 12;
            }
            else
            {
                diam1 = diam1 * 0.0254;
            }

            // Ag.WindowState = FormWindowState.Minimized;
            _AGEN_mainform.tpage_processing.Show();

            try
            {
                _AGEN_mainform.dt_prof = new System.Data.DataTable();

                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_MMid, typeof(string));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station_eq, typeof(double));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev, typeof(string));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Type, typeof(string));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev1, typeof(double));
                _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev2, typeof(double));


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {



                    set_enable_false();

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Polyline3d poly3d = null;
                        double sta_end = -1;
                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        sta_end = poly2d.Length;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            sta_end = poly3d.Length;
                        }



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



                        //double Z = 0;

                        LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        Polyline poly_start = new Polyline();
                        poly_start.AddVertexAt(0, poly2d.GetPoint2dAt(0), 0, 0, 0);
                        poly_start.AddVertexAt(1, new Point2d(poly2d.GetPoint2dAt(0).X + 1000, poly2d.GetPoint2dAt(0).Y), 0, 0, 0);

                        double bear1 = Functions.GET_Bearing_rad(poly2d.GetPoint2dAt(1).X, poly2d.GetPoint2dAt(1).Y, poly2d.GetPoint2dAt(0).X, poly2d.GetPoint2dAt(0).Y);
                        poly_start.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly_start.StartPoint));

                        Polyline poly_end = new Polyline();
                        poly_end.AddVertexAt(0, poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 1), 0, 0, 0);
                        poly_end.AddVertexAt(1, new Point2d(poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 1).X + 1000, poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 1).Y), 0, 0, 0);

                        bear1 = Functions.GET_Bearing_rad(poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 2).X,
                                                                                       poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 2).Y,
                                                                                           poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 1).X,
                                                                                               poly2d.GetPoint2dAt(poly2d.NumberOfVertices - 1).Y);

                        poly_end.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly_end.StartPoint));


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        List<string> lista_field = new List<string>();
                        lista_field.Add("elevation");
                        lista_field.Add("elev");
                        lista_field.Add("elev.");
                        lista_field.Add("z");
                        lista_field.Add("z.");
                        lista_field.Add("el");
                        lista_field.Add("el.");


                        double sta0 = 0;
                        double elev0 = 0;
                        double dist0 = 100000;
                        double elev_at_dist0 = 0;

                        double stan = 0;
                        double elevn = 0;
                        double distn = 100000;
                        double elev_at_distn = 0;




                        foreach (ObjectId ObjID in BTrecord)
                        {


                            Entity Ent_intersection = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;

                            if (Ent_intersection != null)
                            {
                                LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Ent_intersection.Layer], OpenMode.ForRead) as LayerTableRecord;

                                if (Ent_intersection is Curve && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
                                {
                                    ObjID_FAIL = ObjID;
                                    Polyline pstart = new Polyline();
                                    pstart = poly_start.Clone() as Polyline;

                                    Polyline pend = new Polyline();
                                    pend = poly_end.Clone() as Polyline;

                                    Curve Curba_int = Ent_intersection as Curve;

                                    if (Curba_int is Polyline)
                                    {
                                        Polyline Poly2 = (Polyline)Curba_int;
                                        poly2d.Elevation = Poly2.Elevation;
                                        pstart.Elevation = Poly2.Elevation;
                                        pend.Elevation = Poly2.Elevation;


                                    }
                                    if (Curba_int is Line)
                                    {
                                        Line Line2 = (Line)Curba_int;
                                        poly2d.Elevation = Line2.StartPoint.Z;

                                        pstart.Elevation = Line2.StartPoint.Z;
                                        pend.Elevation = Line2.StartPoint.Z;
                                    }

                                    if (Curba_int is Arc)
                                    {
                                        Arc Arc2 = (Arc)Curba_int;
                                        poly2d.Elevation = Arc2.Center.Z;

                                        pstart.Elevation = Arc2.Center.Z;
                                        pend.Elevation = Arc2.Center.Z;
                                    }

                                    if (Curba_int is Ellipse)
                                    {
                                        Ellipse Ellipsa2 = (Ellipse)Curba_int;
                                        poly2d.Elevation = Ellipsa2.Center.Z;

                                        pstart.Elevation = Ellipsa2.Center.Z;
                                        pend.Elevation = Ellipsa2.Center.Z;
                                    }

                                    if (Curba_int is Spline)
                                    {
                                        Spline spl = Curba_int as Spline;
                                        poly2d.Elevation = spl.StartPoint.Z;


                                        pstart.Elevation = spl.StartPoint.Z;
                                        pend.Elevation = spl.StartPoint.Z;
                                    }


                                    Polyline3d P3 = null;
                                    Polyline P2 = null;
                                    if (Curba_int is Polyline3d)
                                    {
                                        P3 = Curba_int as Polyline3d;
                                        P2 = Functions.Build_2dpoly_from_3d(P3);
                                        P2.Elevation = poly2d.Elevation;
                                        Curba_int = P2;

                                    }

                                    if (Curba_int is Circle)
                                    {
                                        Circle C2 = (Circle)Curba_int;
                                        poly2d.Elevation = C2.Center.Z;

                                        pstart.Elevation = C2.Center.Z;
                                        pend.Elevation = C2.Center.Z;
                                    }

                                    Point3dCollection Col_int = new Point3dCollection();
                                    Col_int = Functions.Intersect_on_both_operands(Curba_int, poly2d);

                                    if (Col_int.Count > 0)
                                    {
                                        for (int index = 0; index < Col_int.Count; ++index)
                                        {
                                            _AGEN_mainform.dt_prof.Rows.Add();
                                            bool has_od = Functions.add_object_data_value_to_datatable(_AGEN_mainform.dt_prof, _AGEN_mainform.Col_Elev, lista_field, Tables1, Ent_intersection.ObjectId, true);

                                            Point3d Point_on_poly2d = new Point3d();
                                            Point3d Point_on_poly = new Point3d();
                                            double Station_grid = 0;

                                            Point_on_poly2d = poly2d.GetClosestPointTo(Col_int[index], Vector3d.ZAxis, true);
                                            double Param2d = poly2d.GetParameterAtPoint(Point_on_poly2d);

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                if (Math.Round(poly2d.GetDistanceAtParameter(Param2d), 4) == Math.Round(poly2d.Length, 4))
                                                {
                                                    Param2d = poly3d.EndParam;
                                                }

                                                Station_grid = poly3d.GetDistanceAtParameter(Param2d);
                                                Point_on_poly = poly3d.GetPointAtDist(Station_grid);
                                            }
                                            else
                                            {

                                                Station_grid = poly2d.GetDistanceAtParameter(Param2d);
                                                Point_on_poly = poly2d.GetPointAtDist(Station_grid);
                                            }
                                            _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] = Station_grid;
                                            if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                            {
                                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station_eq] = Functions.Station_equation_ofV2(Station_grid, _AGEN_mainform.dt_station_equation);
                                            }
                                            _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Type] = "On grade";

                                            if (has_od == false)
                                            {
                                                double elevation_from_geometry = poly2d.Elevation;

                                                if (P2 != null && P3 != null)
                                                {
                                                    Point3d Point_on_p2 = P2.GetClosestPointTo(Col_int[index], Vector3d.ZAxis, false);
                                                    double param2 = P2.GetParameterAtPoint(Point_on_p2);
                                                    elevation_from_geometry = P3.GetPointAtParameter(param2).Z;
                                                }


                                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] = Convert.ToString(elevation_from_geometry);
                                            }
                                        }
                                    }

                                    Point3dCollection colint_start = Functions.Intersect_on_both_operands(pstart, Curba_int);
                                    if (colint_start.Count > 0)
                                    {
                                        for (int j = 0; j < colint_start.Count; ++j)
                                        {
                                            double dist = pstart.GetDistAtPoint(pstart.GetClosestPointTo(colint_start[j], Vector3d.ZAxis, false));
                                            if (dist < dist0)
                                            {
                                                dist0 = dist;
                                                elev_at_dist0 = Functions.read_elevation_from_object_data_value(lista_field, Tables1, Ent_intersection.ObjectId);
                                                if (elev_at_dist0 == 0)
                                                {
                                                    elev_at_dist0 = pstart.Elevation;
                                                }
                                            }
                                        }
                                    }

                                    Point3dCollection colint_end = Functions.Intersect_on_both_operands(pend, Curba_int);
                                    if (colint_end.Count > 0)
                                    {
                                        for (int j = 0; j < colint_end.Count; ++j)
                                        {
                                            double dist = pend.GetDistAtPoint(pend.GetClosestPointTo(colint_end[j], Vector3d.ZAxis, false));
                                            if (dist < distn)
                                            {
                                                distn = dist;
                                                elev_at_distn = Functions.read_elevation_from_object_data_value(lista_field, Tables1, Ent_intersection.ObjectId);
                                                if (elev_at_distn == 0)
                                                {
                                                    elev_at_distn = pend.Elevation;
                                                }
                                            }
                                        }
                                    }

                                }
                                ObjID_FAIL = ObjectId.Null;
                            }
                        }

                        _AGEN_mainform.dt_prof = Functions.Sort_data_table(_AGEN_mainform.dt_prof, _AGEN_mainform.Col_station);

                        if (_AGEN_mainform.dt_prof.Rows[0][_AGEN_mainform.Col_Elev] != DBNull.Value)
                        {
                            string val0 = Convert.ToString(_AGEN_mainform.dt_prof.Rows[0][_AGEN_mainform.Col_Elev]);
                            val0 = Functions.extrage_numar_din_text_de_la_inceputul_textului(val0);
                            if (Functions.IsNumeric(val0) == true)
                            {
                                elev0 = Convert.ToDouble(val0);
                            }
                        }


                        if (_AGEN_mainform.dt_prof.Rows[0][_AGEN_mainform.Col_station] != DBNull.Value)
                        {
                            sta0 = Convert.ToDouble(_AGEN_mainform.dt_prof.Rows[0][_AGEN_mainform.Col_station]);
                        }

                        if (_AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] != DBNull.Value)
                        {
                            string val0 = Convert.ToString(_AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev]);
                            val0 = Functions.extrage_numar_din_text_de_la_inceputul_textului(val0);
                            if (Functions.IsNumeric(val0) == true)
                            {
                                elevn = Convert.ToDouble(val0);
                            }
                        }

                        if (_AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] != DBNull.Value)
                        {
                            stan = Convert.ToDouble(_AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station]);
                        }

                        if (_AGEN_mainform.dt_prof != null && _AGEN_mainform.dt_prof.Rows.Count > 0)
                        {
                            if (dist0 < 1000)
                            {
                                double delta_calc = Math.Abs(elev_at_dist0 - elev0) * sta0 / (dist0 + sta0);
                                double elev2 = elev0;

                                if (elev0 < elev_at_dist0)
                                {
                                    elev2 = elev0 + delta_calc;
                                }

                                if (elev0 > elev_at_dist0)
                                {
                                    elev2 = elev0 - delta_calc;
                                }

                                System.Data.DataRow row0 = _AGEN_mainform.dt_prof.NewRow();
                                row0[_AGEN_mainform.Col_station] = 0;
                                row0[_AGEN_mainform.Col_Elev] = Convert.ToString(elev2);
                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    row0[_AGEN_mainform.Col_station_eq] = Functions.Station_equation_ofV2(0, _AGEN_mainform.dt_station_equation);
                                }
                                row0[_AGEN_mainform.Col_Type] = "On grade";
                                _AGEN_mainform.dt_prof.Rows.InsertAt(row0, 0);
                            }



                            if (distn < 1000)
                            {
                                double delta_calc = Math.Abs(elev_at_distn - elevn) * Math.Abs(stan - sta_end) / (distn + sta_end - stan);
                                double elev2 = elevn;

                                if (elevn < elev_at_distn)
                                {
                                    elev2 = elevn + delta_calc;
                                }
                                if (elevn > elev_at_distn)
                                {
                                    elev2 = elevn - delta_calc;
                                }
                                System.Data.DataRow row0 = _AGEN_mainform.dt_prof.NewRow();
                                row0[_AGEN_mainform.Col_station] = sta_end;
                                row0[_AGEN_mainform.Col_Elev] = Convert.ToString(elev2);
                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    row0[_AGEN_mainform.Col_station_eq] = Functions.Station_equation_ofV2(sta_end, _AGEN_mainform.dt_station_equation);
                                }
                                row0[_AGEN_mainform.Col_Type] = "On grade";
                                _AGEN_mainform.dt_prof.Rows.InsertAt(row0, _AGEN_mainform.dt_prof.Rows.Count);
                            }
                        }
                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                        Trans1.Commit();
                    }
                    string fisier_prof = ProjF + _AGEN_mainform.prof_excel_name;
                    Functions.create_backup(fisier_prof);
                    Populate_profile_excel_file(fisier_prof);
                    MessageBox.Show("Done");
                }
                ThisDrawing.Editor.WriteMessage("\nCommand:");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\nsee blue contour line!");

                if (ObjID_FAIL != ObjectId.Null)
                {


                    try
                    {

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                Entity ent1 = Trans1.GetObject(ObjID_FAIL, OpenMode.ForWrite) as Entity;
                                if (ent1 != null)
                                {
                                    ent1.ColorIndex = 5;
                                }

                                Trans1.Commit();
                            }
                        }
                    }
                    catch (System.Exception ex1)
                    {
                        MessageBox.Show(ex1.Message);
                    }



                }

            }
            set_enable_true();
            _AGEN_mainform.tpage_processing.Hide();
        }


        private void read_profile_to_excel_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();
            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.prof_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.prof_excel_name + " file");
                return;
            }



            this.MdiParent.WindowState = FormWindowState.Minimized;

            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }

            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                this.MdiParent.WindowState = FormWindowState.Normal;
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

                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hor;
                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezh = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                Prompt_rezh.MessageForAdding = "\nSelect a known vertical line (STATION) and the label for it:";
                Prompt_rezh.SingleOnly = false;
                Rezultat_hor = ThisDrawing.Editor.GetSelection(Prompt_rezh);

                if (Rezultat_hor.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }


                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_hexag = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\nSpecify horizontal exaggeration:");
                Prompt_hexag.AllowNegative = false;
                Prompt_hexag.AllowZero = false;
                Prompt_hexag.UseDefaultValue = true;
                Prompt_hexag.DefaultValue = 1.0;

                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_hexag = ThisDrawing.Editor.GetDouble(Prompt_hexag);
                if (Rezultat_hexag.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }

                if (Rezultat_hor.Value.Count != 2)
                {
                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_hor.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }

                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ver;
                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rezv = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                Prompt_rezv.MessageForAdding = "\nSelect a known horizontal line (ELEVATION) and the label for it:";
                Prompt_rezv.SingleOnly = false;
                Rezultat_ver = ThisDrawing.Editor.GetSelection(Prompt_rezv);

                if (Rezultat_ver.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }

                if (Rezultat_ver.Value.Count != 2)
                {
                    MessageBox.Show("the selection has to contain 2 objects. You have selected only" + Rezultat_ver.Value.Count.ToString() + " object(s).\r\nOperation aborted");
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }


                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_vexag = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\nSpecify vertical exaggeration:");
                Prompt_vexag.AllowNegative = false;
                Prompt_vexag.AllowZero = false;
                Prompt_vexag.UseDefaultValue = true;
                Prompt_vexag.DefaultValue = 1.0;

                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_vexag = ThisDrawing.Editor.GetDouble(Prompt_vexag);
                if (Rezultat_vexag.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
                    set_enable_true();
                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                    return;
                }


                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly;
                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                Prompt_poly.MessageForAdding = "\nselect profile polyline:";
                Prompt_poly.SingleOnly = true;
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


                        Entity Ent0 = Trans1.GetObject(Rezultat_poly.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                        if ((Ent0 is Polyline) == false)
                        {
                            MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent0.GetType().ToString() + "\r\nOperation aborted");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Double x0 = -123.1234567;
                        Double y0 = -123.1234567;
                        Double sta0 = -123.1234567;
                        Double el0 = -123.1234567;
                        double hexag = Rezultat_hexag.Value;
                        double vexag = Rezultat_vexag.Value;

                        Entity Ent1 = Trans1.GetObject(Rezultat_hor.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                        Entity Ent2 = Trans1.GetObject(Rezultat_hor.Value[1].ObjectId, OpenMode.ForRead) as Entity;

                        Entity Ent3 = Trans1.GetObject(Rezultat_ver.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                        Entity Ent4 = Trans1.GetObject(Rezultat_ver.Value[1].ObjectId, OpenMode.ForRead) as Entity;


                        if (((Ent1 is Polyline || Ent1 is Line) & (Ent2 is MText || Ent2 is DBText)) || ((Ent2 is Polyline || Ent2 is Line) & (Ent1 is MText || Ent1 is DBText)))
                        {

                            if (Ent1 is Polyline)
                            {
                                Polyline P1 = Ent1 as Polyline;
                                if (P1 != null)
                                {
                                    double x1 = P1.StartPoint.X;
                                    double x2 = P1.EndPoint.X;
                                    if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                    {
                                        x0 = x1;

                                    }


                                }

                            }

                            if (Ent1 is Line)
                            {
                                Line L1 = Ent1 as Line;
                                if (L1 != null)
                                {
                                    double x1 = L1.StartPoint.X;
                                    double x2 = L1.EndPoint.X;
                                    if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                    {
                                        x0 = x1;

                                    }


                                }

                            }

                            if (Ent1 is MText)
                            {
                                MText M1 = Ent1 as MText;
                                if (M1 != null)
                                {
                                    string Continut = M1.Text.Replace("+", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        sta0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                            if (Ent1 is DBText)
                            {
                                DBText T1 = Ent1 as DBText;
                                if (T1 != null)
                                {
                                    string Continut = T1.TextString.Replace("+", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        sta0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }


                            if (Ent2 is Polyline)
                            {
                                Polyline P1 = Ent2 as Polyline;
                                if (P1 != null)
                                {
                                    double x1 = P1.StartPoint.X;
                                    double x2 = P1.EndPoint.X;
                                    if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                    {
                                        x0 = x1;

                                    }


                                }

                            }

                            if (Ent2 is Line)
                            {
                                Line L1 = Ent2 as Line;
                                if (L1 != null)
                                {
                                    double x1 = L1.StartPoint.X;
                                    double x2 = L1.EndPoint.X;
                                    if (Math.Round(x1, 2) == Math.Round(x2, 2))
                                    {
                                        x0 = x1;

                                    }


                                }

                            }

                            if (Ent2 is MText)
                            {
                                MText M1 = Ent2 as MText;
                                if (M1 != null)
                                {
                                    string Continut = M1.Text.Replace("+", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        sta0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                            if (Ent2 is DBText)
                            {
                                DBText T1 = Ent2 as DBText;
                                if (T1 != null)
                                {
                                    string Continut = T1.TextString.Replace("+", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        sta0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                        }

                        if (((Ent3 is Polyline || Ent3 is Line) & (Ent4 is MText || Ent4 is DBText)) || ((Ent4 is Polyline || Ent4 is Line) & (Ent3 is MText || Ent3 is DBText)))
                        {

                            if (Ent3 is Polyline)
                            {
                                Polyline P1 = Ent3 as Polyline;
                                if (P1 != null)
                                {
                                    double y1 = P1.StartPoint.Y;
                                    double y2 = P1.EndPoint.Y;
                                    if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                    {
                                        y0 = y1;

                                    }


                                }

                            }

                            if (Ent3 is Line)
                            {
                                Line L1 = Ent3 as Line;
                                if (L1 != null)
                                {
                                    double y1 = L1.StartPoint.Y;
                                    double y2 = L1.EndPoint.Y;
                                    if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                    {
                                        y0 = y1;

                                    }


                                }

                            }

                            if (Ent3 is MText)
                            {
                                MText M1 = Ent3 as MText;
                                if (M1 != null)
                                {
                                    string Continut = M1.Text.Replace("'", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        el0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                            if (Ent3 is DBText)
                            {
                                DBText T1 = Ent3 as DBText;
                                if (T1 != null)
                                {
                                    string Continut = T1.TextString.Replace("'", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        el0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }


                            if (Ent4 is Polyline)
                            {
                                Polyline P1 = Ent4 as Polyline;
                                if (P1 != null)
                                {
                                    double y1 = P1.StartPoint.Y;
                                    double y2 = P1.EndPoint.Y;
                                    if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                    {
                                        y0 = y1;

                                    }


                                }

                            }

                            if (Ent4 is Line)
                            {
                                Line L1 = Ent4 as Line;
                                if (L1 != null)
                                {
                                    double y1 = L1.StartPoint.Y;
                                    double y2 = L1.EndPoint.Y;
                                    if (Math.Round(y1, 2) == Math.Round(y2, 2))
                                    {
                                        y0 = y1;

                                    }


                                }

                            }

                            if (Ent4 is MText)
                            {
                                MText M1 = Ent4 as MText;
                                if (M1 != null)
                                {
                                    string Continut = M1.Text.Replace("'", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        el0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                            if (Ent4 is DBText)
                            {
                                DBText T1 = Ent4 as DBText;
                                if (T1 != null)
                                {
                                    string Continut = T1.TextString.Replace("'", "");
                                    if (Functions.IsNumeric(Continut) == true)
                                    {
                                        el0 = Convert.ToDouble(Continut);

                                    }


                                }

                            }

                        }

                        double depth1 = 0;
                        double diam1 = 0;

                        if (Functions.IsNumeric(textBox_cover.Text) == true)
                        {
                            depth1 = Convert.ToDouble(textBox_cover.Text);
                        }

                        if (Functions.IsNumeric(textBox_pipe_diam.Text) == true)
                        {
                            diam1 = Convert.ToDouble(textBox_pipe_diam.Text);
                        }

                        if (_AGEN_mainform.COUNTRY == "USA")
                        {
                            diam1 = diam1 / 12;
                        }
                        else
                        {
                            diam1 = diam1 * 0.0254;
                        }

                        Polyline poly1 = Ent0 as Polyline;
                        if (poly1 != null && x0 != -123.1234567 && y0 != -123.1234567 && sta0 != -123.1234567 && el0 != -123.1234567)
                        {
                            _AGEN_mainform.dt_prof = new System.Data.DataTable();

                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_MMid, typeof(string));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station_eq, typeof(double));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev, typeof(double));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Type, typeof(string));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev1, typeof(double));
                            _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev2, typeof(double));


                            for (int i = 0; i < poly1.NumberOfVertices; ++i)
                            {
                                double x = poly1.GetPointAtParameter(i).X;
                                double y = poly1.GetPointAtParameter(i).Y;

                                double Sta = sta0 + (x - x0) / hexag;
                                double El = el0 + (y - y0) / vexag;
                                _AGEN_mainform.dt_prof.Rows.Add();
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] = Sta;
                                if (_AGEN_mainform.dt_station_equation != null)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {
                                        _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station_eq] = Functions.Station_equation_of(Sta, _AGEN_mainform.dt_station_equation);
                                    }
                                }

                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Type] = "On grade";
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] = El;

                                if (diam1 > 0 && depth1 != 0)
                                {
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev1] = El - depth1;
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev2] = El - depth1 - diam1;
                                }
                            }

                            _AGEN_mainform.dt_prof = Functions.Sort_data_table(_AGEN_mainform.dt_prof, _AGEN_mainform.Col_station);

                            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                            {
                                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }


                                string fisier_prof = ProjF + _AGEN_mainform.prof_excel_name;

                                Populate_profile_excel_file(fisier_prof);
                            }

                        }
                        else
                        {
                            MessageBox.Show("you did not selected the proper entities");
                        }


                        Trans1.Dispose();
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


        private void Populate_profile_excel_file(string File1)
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
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_prof, _AGEN_mainform.Start_row_graph_profile, "General");
                    Functions.Create_header_graph_profile(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
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

        private void button_show_profile_draw_Click(object sender, EventArgs e)
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

            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_profdraw.Show();

        }

        private void label_scan_profile_data_Click(object sender, EventArgs e)
        {

        }



        private void button_generate_profile3D_xls_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.prof_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.prof_excel_name + " file");
                return;
            }


            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            set_enable_false();
            try
            {
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
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);


                    if (_AGEN_mainform.dt_centerline != null && _AGEN_mainform.dt_centerline.Rows.Count > 0)
                    {
                        _AGEN_mainform.dt_prof = new System.Data.DataTable();
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_MMid, typeof(string));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station_eq, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Type, typeof(string));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev1, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev2, typeof(double));

                        double depth1 = 0;
                        double diam1 = 0;

                        if (Functions.IsNumeric(textBox_cover.Text) == true)
                        {
                            depth1 = Convert.ToDouble(textBox_cover.Text);
                        }

                        if (Functions.IsNumeric(textBox_pipe_diam.Text) == true)
                        {
                            diam1 = Convert.ToDouble(textBox_pipe_diam.Text);
                        }

                        if (_AGEN_mainform.COUNTRY == "USA")
                        {
                            diam1 = diam1 / 12;
                        }
                        else
                        {
                            diam1 = diam1 * 0.0254;
                        }

                        for (int i = 0; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_z] != DBNull.Value &&
                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                            {
                                double z = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_z]);
                                double sta = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta]);
                                double ahead1 = -1.23456;

                                if (_AGEN_mainform.dt_centerline.Rows[i][Col_BackSta] != DBNull.Value)
                                {
                                    sta = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_BackSta]);
                                }

                                if (_AGEN_mainform.dt_centerline.Rows[i][Col_AheadSta] != DBNull.Value)
                                {
                                    ahead1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][Col_AheadSta]);
                                }

                                _AGEN_mainform.dt_prof.Rows.Add();
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] = sta;
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] = z;
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Type] = "NG";
                                if (ahead1 != -1.23456)
                                {
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station_eq] = ahead1;
                                }

                                if (diam1 > 0 && depth1 != 0)
                                {
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev1] = z - depth1;
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev2] = z - depth1 - diam1;
                                }

                            }
                            else
                            {
                                set_enable_true();
                                MessageBox.Show("the centerline data file data is not correct");
                                _AGEN_mainform.dt_station_equation = null;
                                _AGEN_mainform.dt_prof = null;
                                return;
                            }

                        }
                        string fisier_prof = ProjF + _AGEN_mainform.prof_excel_name;
                        Functions.create_backup(fisier_prof);
                        Populate_profile_excel_file(fisier_prof);
                    }

                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
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

                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        Polyline3d poly3d = null;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                        }


                        for (int i = 0; i < Rezultat_poly.Value.Count; ++i)
                        {
                            Entity Ent1 = Trans1.GetObject(Rezultat_poly.Value[i].ObjectId, OpenMode.ForRead) as Entity;
                            if (Ent1 != null && (Ent1 is Line || Ent1 is Polyline || Ent1 is Polyline3d || Ent1 is Polyline2d))
                            {
                                Curve line1 = Ent1 as Curve;
                                Point3d p1 = line1.StartPoint;
                                Point3d p2 = line1.EndPoint;


                                Point3d pt1 = poly2d.GetClosestPointTo(p1, Vector3d.ZAxis, false);
                                Point3d pt2 = poly2d.GetClosestPointTo(p2, Vector3d.ZAxis, false);
                                double d1 = poly2d.GetDistAtPoint(pt1);
                                double d2 = poly2d.GetDistAtPoint(pt2);

                                if (_AGEN_mainform.Project_type == "3D")
                                {
                                    double param1 = poly2d.GetParameterAtPoint(pt1);
                                    double param2 = poly2d.GetParameterAtPoint(pt2);


                                    pt1 = poly3d.GetPointAtParameter(param1);
                                    pt2 = poly3d.GetPointAtParameter(param2);
                                    d1 = poly3d.GetDistanceAtParameter(param1);
                                    d2 = poly3d.GetDistanceAtParameter(param2);
                                }


                                if (d1 > d2)
                                {
                                    Point3d t = pt1;
                                    pt1 = pt2;
                                    pt2 = t;

                                    double tt = d1;
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

                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                {
                                    double d1_2d = poly2d.GetDistAtPoint(pt1);
                                    double d2_2d = poly2d.GetDistAtPoint(pt2);
                                    double b1 = -1.23456;
                                    double b2 = -1.23456;
                                    double Sta1 = Functions.get_stationCSF_from_point(poly2d, pt1, d1_2d, _AGEN_mainform.dt_centerline, ref b1);
                                    double Sta2 = Functions.get_stationCSF_from_point(poly2d, pt2, d2_2d, _AGEN_mainform.dt_centerline, ref b2);
                                    dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(Sta1, _AGEN_mainform.round1);
                                    dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(Sta2, _AGEN_mainform.round1);
                                }
                                else
                                {
                                    dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(d1, _AGEN_mainform.round1);
                                    dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(d2, _AGEN_mainform.round1);
                                }
                            }
                        }

                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
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

        private void button_calc_low_high_Click(object sender, EventArgs e)
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

            Functions.Kill_excel();


            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }


            string fisier_prof = ProjFolder + _AGEN_mainform.prof_excel_name;

            if (System.IO.File.Exists(fisier_prof) == false)
            {
                MessageBox.Show("the profile data file does not exist");
                return;
            }


            System.Data.DataTable dt_null = null;
            System.Data.DataTable dt_prof = _AGEN_mainform.tpage_profdraw.Load_existing_profile_graph(fisier_prof, ref dt_null);

            if (dt_prof != null)
            {
                if (dt_prof.Rows.Count > 0)
                {



                    string text_sta1 = textBox_start_sta.Text;
                    string text_sta2 = textBox_end_sta.Text;

                    if (Functions.IsNumeric(text_sta1.Replace("+", "")) == true && Functions.IsNumeric(text_sta2.Replace("+", "")) == true)
                    {
                        double sta1 = Convert.ToDouble(text_sta1.Replace("+", ""));
                        double sta2 = Convert.ToDouble(text_sta2.Replace("+", ""));

                        if (sta1 > sta2)
                        {
                            double t = sta1;
                            sta1 = sta2;
                            sta2 = t;
                        }

                        bool start1 = false;
                        bool end1 = false;

                        double stap = -1.234;
                        double elevp = -1.234;

                        double low_sta = -1.234;
                        double high_sta = -1.234;
                        double low_elev = -1.234;
                        double high_elev = -1.234;
                        double elev_start = -1.234;
                        double elev_end = -1.234;

                        for (int i = 0; i < dt_prof.Rows.Count; ++i)
                        {
                            if (dt_prof.Rows[i]["Station"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_prof.Rows[i]["Station"])) == true &&
                                dt_prof.Rows[i]["Elev"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_prof.Rows[i]["Elev"])) == true)
                            {
                                double sta = Convert.ToDouble(dt_prof.Rows[i]["Station"]);
                                double elev = Convert.ToDouble(dt_prof.Rows[i]["Elev"]);

                                if (sta1 == sta)
                                {
                                    start1 = true;
                                    low_sta = sta1;
                                    high_sta = sta1;
                                    low_elev = elev;
                                    high_elev = elev;
                                    elev_start = elev;

                                }

                                else if (sta2 == sta)
                                {
                                    end1 = true;
                                    if (low_elev > elev)
                                    {
                                        low_elev = elev;
                                        low_sta = sta2;
                                    }
                                    if (high_elev < elev)
                                    {
                                        high_elev = elev;
                                        high_sta = sta2;
                                    }
                                    elev_end = elev;

                                }

                                else if (sta > sta1 && start1 == false)
                                {
                                    start1 = true;
                                    low_sta = sta1;
                                    high_sta = sta1;

                                    double elt = elevp + (elev - elevp) * (sta1 - stap) / (sta - stap);

                                    low_elev = elt;
                                    high_elev = elt;
                                    elev_start = elt;

                                    if (sta2 > sta)
                                    {
                                        if (low_elev > elev)
                                        {
                                            low_elev = elev;
                                            low_sta = sta;
                                        }
                                        if (high_elev < elev)
                                        {
                                            high_elev = elev;
                                            high_sta = sta;
                                        }
                                    }


                                    if (sta > sta2 && end1 == false)
                                    {
                                        end1 = true;


                                        elt = elevp + (elev - elevp) * (sta2 - stap) / (sta - stap);
                                        elev_end = elt;

                                        if (low_elev > elt)
                                        {
                                            low_elev = elt;
                                            low_sta = sta2;
                                        }
                                        if (high_elev < elt)
                                        {
                                            high_elev = elt;
                                            high_sta = sta2;
                                        }

                                    }


                                }

                                else if (sta > sta2 && end1 == false)
                                {
                                    end1 = true;
                                    double elt = elevp + (elev - elevp) * (sta2 - stap) / (sta - stap);
                                    elev_end = elt;

                                    if (low_elev > elt)
                                    {
                                        low_elev = elt;
                                        low_sta = sta2;
                                    }
                                    if (high_elev < elt)
                                    {
                                        high_elev = elt;
                                        high_sta = sta2;
                                    }

                                }

                                else if (sta > sta1 && start1 == true && end1 == false)
                                {

                                    if (low_elev > elev)
                                    {
                                        low_elev = elev;
                                        low_sta = sta;
                                    }
                                    if (high_elev < elev)
                                    {
                                        high_elev = elev;
                                        high_sta = sta;
                                    }


                                }


                                stap = sta;
                                elevp = elev;




                            }

                        }

                        textBox_high_sta.Text = Functions.Get_chainage_from_double(high_sta, _AGEN_mainform.units_of_measurement, 2);
                        textBox_low_sta.Text = Functions.Get_chainage_from_double(low_sta, _AGEN_mainform.units_of_measurement, 2);
                        textBox_high_elev.Text = Functions.Get_String_Rounded(high_elev, 2);
                        textBox_low_elev.Text = Functions.Get_String_Rounded(low_elev, 2);

                        ++rec_no;
                        label_count.Text = Convert.ToString(rec_no);

                        if (dt_hl == null)
                        {
                            dt_hl = new System.Data.DataTable();
                            dt_hl.Columns.Add("Test Section", typeof(int));
                            dt_hl.Columns.Add("Start", typeof(double));
                            dt_hl.Columns.Add("Start_Elev", typeof(double));
                            dt_hl.Columns.Add("End", typeof(double));
                            dt_hl.Columns.Add("End_Elev", typeof(double));
                            dt_hl.Columns.Add("High_Point_Sta", typeof(double));
                            dt_hl.Columns.Add("High_Point_Elev", typeof(double));
                            dt_hl.Columns.Add("Low_Point_Sta", typeof(double));
                            dt_hl.Columns.Add("Low_Point_Elev", typeof(double));

                        }

                        dt_hl.Rows.Add();
                        dt_hl.Rows[dt_hl.Rows.Count - 1][0] = rec_no;
                        dt_hl.Rows[dt_hl.Rows.Count - 1][1] = sta1;
                        dt_hl.Rows[dt_hl.Rows.Count - 1][2] = Math.Round(elev_start, 2);
                        dt_hl.Rows[dt_hl.Rows.Count - 1][3] = sta2;
                        dt_hl.Rows[dt_hl.Rows.Count - 1][4] = Math.Round(elev_end, 2);
                        dt_hl.Rows[dt_hl.Rows.Count - 1][5] = Math.Round(high_sta, 2);
                        dt_hl.Rows[dt_hl.Rows.Count - 1][6] = Math.Round(high_elev, 2);
                        dt_hl.Rows[dt_hl.Rows.Count - 1][7] = Math.Round(low_sta, 2);
                        dt_hl.Rows[dt_hl.Rows.Count - 1][8] = Math.Round(low_elev, 2);


                    }
                    else
                    {
                        MessageBox.Show("not numeric stations specified");
                    }
                }
                else
                {
                    MessageBox.Show("no profile data found");
                }
            }
            else
            {
                MessageBox.Show("no profile data found");
            }

            set_enable_true();
        }

        private void button_export_to_excel_high_low_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_hl);
            dt_hl = null;
            rec_no = 0;
            label_count.Text = Convert.ToString(rec_no);
        }

        private void button_draw_Mleader_Click(object sender, EventArgs e)
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



            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }




            double Hexag = 0;
            if (Functions.IsNumeric(textBox_prof_Hex.Text) == true)
            {
                Hexag = Convert.ToDouble(textBox_prof_Hex.Text);
            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("specify the profile horizontal exxageration");
                return;
            }


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


                        int lr = 1;

                        if (_AGEN_mainform.Left_to_Right == false)
                        {
                            lr = -1;
                        }

                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        double ymin = -1000000;
                        double ymax = 1000000;

                        Polyline Poly2d = new Polyline();



                        List<ObjectId> lista_poly = new List<ObjectId>();
                        List<double> lista_start = new List<double>();
                        List<double> lista_end = new List<double>();

                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        string Agen_profile_band = "Agen_profile_band";
                        string Agen_profile_band_V2 = "Agen_profile_band_V2";
                        string Agen_profile_band_V3 = "Agen_profile_band_V3";

                        #region profile band
                        if (Tables1.IsTableDefined(Agen_profile_band) == true || Tables1.IsTableDefined(Agen_profile_band_V2) == true || Tables1.IsTableDefined(Agen_profile_band_V3) == true)
                        {
                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline poly_ground = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                if (poly_ground != null)
                                {

                                    #region old data table
                                    if (Tables1.IsTableDefined(Agen_profile_band) == true)
                                    {
                                        using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band])
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
                                                        }

                                                        if (start1 != -123.4 && end1 != 123.4)
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
                                    #endregion


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
                                                        string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                        if (segment1 == "not defined") segment1 = "";
                                                        if (start1 != -123.4 && end1 != 123.4 && segm1.ToLower() == segment1.ToLower())
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
                                    if (Tables1.IsTableDefined(Agen_profile_band_V3) == true)
                                    {
                                        using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V3])
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
                                                        string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                        if (segment1 == "not defined") segment1 = "";
                                                        if (start1 != -123.4 && end1 != 123.4 && segm1.ToLower() == segment1.ToLower())
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
                        #endregion



                        Functions.Creaza_layer(_AGEN_mainform.layer_prof_block_labels, 2, true);
                        double Texth = 8;
                        if (Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
                        {
                            Texth = Convert.ToDouble(textBox_overwrite_text_height.Text);
                        }


                        if (Functions.IsNumeric(textBox_low_sta.Text.Replace("+", "")) == true)
                        {
                            double Station1 = Convert.ToDouble(textBox_low_sta.Text.Replace("+", ""));
                            if (Station1 >= 0)
                            {
                                #region profile band work
                                if (lista_start.Count > 0 && lista_start.Count == lista_end.Count && lista_start.Count == lista_poly.Count)
                                {
                                    for (int k = 0; k < lista_poly.Count; ++k)
                                    {
                                        if (lista_poly[k] != null && lista_poly[k] != ObjectId.Null)
                                        {
                                            Poly2d = Trans1.GetObject(lista_poly[k], OpenMode.ForRead) as Polyline;
                                            if (Poly2d != null)
                                            {
                                                double start1 = lista_start[k];
                                                double end1 = lista_end[k];

                                                if (Station1 >= start1 && Station1 <= end1)
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
                                                    double x = Poly2d.StartPoint.X + lr * (Station1 - start1) * Hexag;
                                                    Line line1 = new Line(new Point3d(x, ymin - 10000, Poly2d.Elevation), new Point3d(x, ymax + 10000, Poly2d.Elevation));
                                                    Point3dCollection col1 = Functions.Intersect_on_both_operands(Poly2d, line1);
                                                    Point3d inspt = new Point3d();
                                                    if (col1.Count > 0)
                                                    {
                                                        inspt = col1[0];
                                                    }
                                                    else
                                                    {
                                                        inspt = new Point3d(x, Poly2d.GetPoint2dAt(0).Y, Poly2d.Elevation);
                                                    }
                                                    string descriptie = "LOW POINT\r\nSTA: " + textBox_low_sta.Text + "\r\nELEV: " + textBox_low_elev.Text;

                                                    if (textBox_test_section.Text.Replace(" ", "").Length > 0)
                                                    {
                                                        descriptie = textBox_test_section.Text + "\r\n" + descriptie;
                                                    }
                                                    if (_AGEN_mainform.COUNTRY == "USA") descriptie = descriptie + "'";
                                                    MLeader mleader1 = Functions.creaza_mleader(inspt, descriptie, Texth, 3 * Texth, 7 * Texth, Texth, Texth, Texth);
                                                    mleader1.Layer = _AGEN_mainform.layer_prof_block_labels;
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }


                        }

                        if (Functions.IsNumeric(textBox_high_sta.Text.Replace("+", "")) == true)
                        {
                            double Station1 = Convert.ToDouble(textBox_high_sta.Text.Replace("+", ""));
                            if (Station1 >= 0)
                            {
                                #region profile band work
                                if (lista_start.Count > 0 && lista_start.Count == lista_end.Count && lista_start.Count == lista_poly.Count)
                                {
                                    for (int k = 0; k < lista_poly.Count; ++k)
                                    {
                                        if (lista_poly[k] != null && lista_poly[k] != ObjectId.Null)
                                        {
                                            Poly2d = Trans1.GetObject(lista_poly[k], OpenMode.ForRead) as Polyline;
                                            if (Poly2d != null)
                                            {
                                                double start1 = lista_start[k];
                                                double end1 = lista_end[k];

                                                if (Station1 >= start1 && Station1 <= end1)
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
                                                    double x = Poly2d.StartPoint.X + lr * (Station1 - start1) * Hexag;
                                                    Line line1 = new Line(new Point3d(x, ymin - 10000, Poly2d.Elevation), new Point3d(x, ymax + 10000, Poly2d.Elevation));
                                                    Point3dCollection col1 = Functions.Intersect_on_both_operands(Poly2d, line1);
                                                    Point3d inspt = new Point3d();
                                                    if (col1.Count > 0)
                                                    {
                                                        inspt = col1[0];
                                                    }
                                                    else
                                                    {
                                                        inspt = new Point3d(x, Poly2d.GetPoint2dAt(0).Y, Poly2d.Elevation);
                                                    }
                                                    string descriptie = "HIGH POINT\r\nSTA: " + textBox_high_sta.Text + "\r\nELEV: " + textBox_high_elev.Text;

                                                    if (textBox_test_section.Text.Replace(" ", "").Length > 0)
                                                    {
                                                        descriptie = textBox_test_section.Text + "\r\n" + descriptie;
                                                    }
                                                    if (_AGEN_mainform.COUNTRY == "USA") descriptie = descriptie + "'";
                                                    MLeader mleader1 = Functions.creaza_mleader(inspt, descriptie, Texth, 3 * Texth, 7 * Texth, Texth, Texth, Texth);
                                                    mleader1.Layer = _AGEN_mainform.layer_prof_block_labels;
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion
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


        private void Panel7_Click(object sender, EventArgs e)
        {
            if (panel_high_low.Visible == false)
            {
                panel_high_low.Visible = true;
            }
            else
            {
                panel_high_low.Visible = false;
            }
        }

        private void button_generate_profile2D_xls_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.prof_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.prof_excel_name + " file");
                return;
            }


            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            set_enable_false();
            try
            {
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
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);


                    if (_AGEN_mainform.dt_centerline != null && _AGEN_mainform.dt_centerline.Rows.Count > 0)
                    {
                        _AGEN_mainform.dt_prof = new System.Data.DataTable();
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_MMid, typeof(string));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_station_eq, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Type, typeof(string));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev1, typeof(double));
                        _AGEN_mainform.dt_prof.Columns.Add(_AGEN_mainform.Col_Elev2, typeof(double));

                        double depth1 = 0;
                        double diam1 = 0;

                        if (Functions.IsNumeric(textBox_cover.Text) == true)
                        {
                            depth1 = Convert.ToDouble(textBox_cover.Text);
                        }

                        if (Functions.IsNumeric(textBox_pipe_diam.Text) == true)
                        {
                            diam1 = Convert.ToDouble(textBox_pipe_diam.Text);
                        }

                        if (_AGEN_mainform.COUNTRY == "USA")
                        {
                            diam1 = diam1 / 12;
                        }
                        else
                        {
                            diam1 = diam1 * 0.0254;
                        }

                        double sta0 = 0;

                        if (_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_z] != DBNull.Value)
                        {
                            double z = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_z]);
                            if (_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                            {
                                sta0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[0][_AGEN_mainform.Col_3DSta]);
                            }
                            _AGEN_mainform.dt_prof.Rows.Add();
                            _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] = sta0;
                            _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] = z;
                            _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Type] = "NG";
                        }

                        for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_centerline.Rows[i - 1][_AGEN_mainform.Col_x] != DBNull.Value &&
                                _AGEN_mainform.dt_centerline.Rows[i - 1][_AGEN_mainform.Col_y] != DBNull.Value &&
                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value &&
                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value &&
                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_z] != DBNull.Value)
                            {
                                double x0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][_AGEN_mainform.Col_x]);
                                double y0 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1][_AGEN_mainform.Col_y]);
                                double x = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_x]);
                                double y = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_y]);
                                double z = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_z]);
                                double sta = sta0 + Math.Pow(Math.Pow(x - x0, 2) + Math.Pow(y - y0, 2), 0.5);

                                _AGEN_mainform.dt_prof.Rows.Add();
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] = sta;
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev] = z;
                                _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Type] = "NG";

                                sta0 = sta;

                                if (diam1 > 0 && depth1 != 0)
                                {
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev1] = z - depth1;
                                    _AGEN_mainform.dt_prof.Rows[_AGEN_mainform.dt_prof.Rows.Count - 1][_AGEN_mainform.Col_Elev2] = z - depth1 - diam1;
                                }

                            }
                            else
                            {
                                set_enable_true();
                                MessageBox.Show("the centerline data file data is not correct");
                                _AGEN_mainform.dt_station_equation = null;
                                _AGEN_mainform.dt_prof = null;
                                return;
                            }

                        }
                        string fisier_prof = ProjF + _AGEN_mainform.prof_excel_name;
                        Functions.create_backup(fisier_prof);
                        Populate_profile_excel_file(fisier_prof);
                    }

                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }
    }
}
