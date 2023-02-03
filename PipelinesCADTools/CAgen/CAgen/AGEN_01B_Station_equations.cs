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
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    /// <summary>
    /// 
    /// </summary>
    public partial class AGEN_station_equations : Form
    {
        _AGEN_mainform Ag = null;
        System.Data.DataTable dt_display = null;

        public AGEN_station_equations()
        {
            InitializeComponent();

        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_station_equations);
            lista_butoane.Add(button_redefine_begin_station);
            lista_butoane.Add(button_add_prefix_r2);
            lista_butoane.Add(button_Station_eq_pick_sta);
            lista_butoane.Add(button_transf_to_excel);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_calc_station_equations);
            lista_butoane.Add(button_redefine_begin_station);
            lista_butoane.Add(button_add_prefix_r2);
            lista_butoane.Add(button_Station_eq_pick_sta);
            lista_butoane.Add(button_transf_to_excel);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void button_calc_station_equations_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;
            

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.cl_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.cl_excel_name + " file");
                return;
            }

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjF) == false)
            {
                MessageBox.Show("No project Loaded");
                return;
            }

            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }

            string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;


            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            string fisier_cs = ProjF + _AGEN_mainform.crossing_excel_name;
            string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;


            _AGEN_mainform.tpage_processing.Show();

            Ag.WindowState = FormWindowState.Minimized;


            set_enable_false();

            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect Reroute centerlines:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            Ag.WindowState = FormWindowState.Normal;
                            _AGEN_mainform.tpage_processing.Hide();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            return;
                        }

                        Ag.WindowState = FormWindowState.Normal;

                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                        Functions.create_backup(fisier_cl);

                        if (System.IO.File.Exists(fisier_cs) == true)
                        {
                            Functions.create_backup(fisier_cs);
                            _AGEN_mainform.Data_Table_crossings = _AGEN_mainform.tpage_crossing_scan.Load_existing_crossing(fisier_cs);
                        }
                        else
                        {
                            _AGEN_mainform.Data_Table_crossings = Functions.Creaza_crossing_datatable_structure();
                        }

                        if (System.IO.File.Exists(fisier_prop) == true)
                        {
                            Functions.create_backup(fisier_prop);
                            _AGEN_mainform.Data_Table_property = _AGEN_mainform.tpage_setup.Load_existing_property(fisier_prop);
                        }
                        else
                        {
                            _AGEN_mainform.Data_Table_property = Functions.Creaza_property_datatable_structure();
                        }


                        Polyline old_poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);




                        List<Polyline> lista1 = new List<Polyline>();
                        List<ObjectId> lista_reroutes = new List<ObjectId>();
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Polyline poly_reroute = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;

                            if (poly_reroute != null)
                            {
                                lista1.Add(poly_reroute);
                            }
                        }



                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        if (lista1.Count > 0)
                        {
                            System.Data.DataTable dt_temp_seq = Functions.Creaza_station_equation_datatable_structure();
                            dt_temp_seq.Columns.Add("measured", typeof(double));


                            for (int i = 0; i < lista1.Count; ++i)
                            {
                                Polyline poly2d_int = Trans1.GetObject(lista1[i].ObjectId, OpenMode.ForWrite) as Polyline;

                                if (poly2d_int != null)
                                {
                                    double Orig_elev = poly2d_int.Elevation;

                                    poly2d_int.Elevation = old_poly2D.Elevation;

                                    Point3d pt_on_poly1 = old_poly2D.GetClosestPointTo(poly2d_int.StartPoint, Vector3d.ZAxis, false);

                                    double dist1 = pt_on_poly1.DistanceTo(poly2d_int.StartPoint);

                                    Point3d pt_on_poly2 = old_poly2D.GetClosestPointTo(poly2d_int.EndPoint, Vector3d.ZAxis, false);

                                    double dist2 = pt_on_poly2.DistanceTo(poly2d_int.EndPoint);

                                    if (dist1 < 0.01 && dist2 < 0.01)
                                    {
                                        double param1 = old_poly2D.GetParameterAtPoint(pt_on_poly1);
                                        double param2 = old_poly2D.GetParameterAtPoint(pt_on_poly2);

                                        if (param1 > param2)
                                        {
                                            double temp = param1;
                                            param1 = param2;
                                            param2 = temp;
                                        }



                                        Point3d Pt_start = old_poly2D.GetPointAtParameter(param1);
                                        Point3d Pt_End = old_poly2D.GetPointAtParameter(param2);

                                        dt_temp_seq.Rows.Add();
                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute Start X"] = Pt_start.X;
                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute Start Y"] = Pt_start.Y;
                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute Start Z"] = 0;

                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute End X"] = Pt_End.X;
                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute End Y"] = Pt_End.Y;
                                        dt_temp_seq.Rows[dt_temp_seq.Rows.Count - 1]["Reroute End Z"] = 0;

                                        poly2d_int.Elevation = Orig_elev;

                                        lista_reroutes.Add(poly2d_int.ObjectId);
                                    }
                                    else
                                    {
                                        MessageBox.Show("the reroute polyline doesn't touch the cl polyline");
                                        lista_reroutes.Clear();
                                        Functions.zoom_to_object(poly2d_int.ObjectId);

                                        Trans1.Commit();
                                        _AGEN_mainform.tpage_processing.Hide();
                                        set_enable_true();
                                        return;

                                    }
                                }
                            }


                            if (dt_temp_seq.Rows.Count > 0)
                            {
                                Polyline poly2d = _AGEN_mainform.tpage_setup.create_new_centerline(lista_reroutes); //NEW POLY2D!
                                Polyline3d poly3d = null;
                                if (_AGEN_mainform.Project_type == "3D")
                                {
                                    poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                }


                                if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false)
                                {
                                    _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));
                                }

                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                    {
                                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                        {
                                            double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                            double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);

                                            Point3d pt_on_2d = old_poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                            double param1 = old_poly2D.GetParameterAtPoint(pt_on_2d);
                                            double eq_meas = old_poly2D.GetDistanceAtParameter(param1);
                                            _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;


                                        }
                                    }
                                }

                                for (int i = 0; i < dt_temp_seq.Rows.Count; ++i)
                                {
                                    double x2 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute End X"]);
                                    double y2 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute End Y"]);
                                    Point3d ptn2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);
                                    double stan2 = poly2d.GetDistAtPoint(ptn2);
                                    dt_temp_seq.Rows[i]["measured"] = stan2;
                                }

                                dt_temp_seq = Functions.Sort_data_table(dt_temp_seq, "measured");


                                for (int i = 0; i < dt_temp_seq.Rows.Count; ++i)
                                {
                                    double x1 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute Start X"]);
                                    double y1 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute Start Y"]);
                                    double x2 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute End X"]);
                                    double y2 = Convert.ToDouble(dt_temp_seq.Rows[i]["Reroute End Y"]);

                                    Point3d pt1 = old_poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                    double sta1 = old_poly2D.GetDistAtPoint(pt1);
                                    double eq1 = Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation);

                                    Point3d pt2 = old_poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);
                                    double sta2 = old_poly2D.GetDistAtPoint(pt2);
                                    double SA = Functions.Station_equation_ofV2(sta2, _AGEN_mainform.dt_station_equation);

                                    Point3d ptn1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                    Point3d ptn2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                    double stan1 = poly2d.GetDistAtPoint(ptn1);
                                    double stan2 = poly2d.GetDistAtPoint(ptn2);

                                    double SB = eq1 + stan2 - stan1;

                                    dt_temp_seq.Rows[i]["Station Back"] = SB;
                                    dt_temp_seq.Rows[i]["Station Ahead"] = SA;

                                }


                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    for (int i = _AGEN_mainform.dt_station_equation.Rows.Count - 1; i >= 0; --i)
                                    {
                                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                        {
                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start X"]);
                                            double y1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Y"]);
                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);

                                            Point3d pt1 = new Point3d(x1, y1, 0);
                                            Point3d pt2 = new Point3d(x2, y2, 0);

                                            Point3d ptn1 = poly2d.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                                            Point3d ptn2 = poly2d.GetClosestPointTo(pt2, Vector3d.ZAxis, false);

                                            double param2 = poly2d.GetParameterAtPoint(ptn2);
                                            double eq2 = poly2d.GetDistanceAtParameter(param2);
                                            _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq2;

                                            double SB = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"]);

                                            double dist1 = pt1.DistanceTo(ptn1);
                                            double dist2 = pt2.DistanceTo(ptn2);

                                            if (dist2 > 0.01)
                                            {
                                                _AGEN_mainform.dt_station_equation.Rows[i].Delete();
                                            }
                                            else if (dist1 > 0.01 && dist2 <= 0.01)
                                            {
                                                Point3d pts1 = old_poly2D.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                                                double stax = old_poly2D.GetDistAtPoint(pts1);

                                                for (int j = 0; j < dt_temp_seq.Rows.Count; ++j)
                                                {
                                                    double xp1 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute Start X"]);
                                                    double yp1 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute Start Y"]);

                                                    double xp2 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute End X"]);
                                                    double yp2 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute End Y"]);


                                                    Point3d ptp1 = old_poly2D.GetClosestPointTo(new Point3d(xp1, yp1, 0), Vector3d.ZAxis, false);
                                                    double stap1 = old_poly2D.GetDistAtPoint(ptp1);

                                                    Point3d ptp2 = old_poly2D.GetClosestPointTo(new Point3d(xp2, yp2, 0), Vector3d.ZAxis, false);
                                                    double stap2 = old_poly2D.GetDistAtPoint(ptp2);

                                                    if (stax >= stap1 && stax <= stap2)
                                                    {
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start X"] = xp2;
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Y"] = yp2;
                                                        j = dt_temp_seq.Rows.Count;
                                                    }
                                                }
                                            }
                                            else
                                            {


                                                Point3d pts1 = old_poly2D.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                                                double stax1 = old_poly2D.GetDistAtPoint(pts1);

                                                Point3d pts2 = old_poly2D.GetClosestPointTo(pt2, Vector3d.ZAxis, false);
                                                double stax2 = old_poly2D.GetDistAtPoint(pts2);


                                                for (int j = 0; j < dt_temp_seq.Rows.Count; ++j)
                                                {
                                                    double xp1 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute Start X"]);
                                                    double yp1 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute Start Y"]);

                                                    double xp2 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute End X"]);
                                                    double yp2 = Convert.ToDouble(dt_temp_seq.Rows[j]["Reroute End Y"]);


                                                    Point3d ptp1 = old_poly2D.GetClosestPointTo(new Point3d(xp1, yp1, 0), Vector3d.ZAxis, false);
                                                    double stap1 = old_poly2D.GetDistAtPoint(ptp1);

                                                    Point3d ptp2 = old_poly2D.GetClosestPointTo(new Point3d(xp2, yp2, 0), Vector3d.ZAxis, false);
                                                    double stap2 = old_poly2D.GetDistAtPoint(ptp2);

                                                    if (stax1 >= stap1 && stax1 <= stap2 && stax2 >= stap1 && stax2 <= stap2)
                                                    {
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start X"] = xp2;
                                                        _AGEN_mainform.dt_station_equation.Rows[i]["Reroute Start Y"] = yp2;

                                                        j = dt_temp_seq.Rows.Count;
                                                    }
                                                }



                                            }
                                        }
                                    }

                                }

                                for (int i = 0; i < dt_temp_seq.Rows.Count; ++i)
                                {
                                    System.Data.DataRow row1 = dt_temp_seq.Rows[i];
                                    _AGEN_mainform.dt_station_equation.ImportRow(row1);
                                }

                                _AGEN_mainform.dt_station_equation = Functions.Sort_data_table(_AGEN_mainform.dt_station_equation, "measured");







                                #region crossing xlxs   
                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                {
                                    if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                                    {
                                        for (int i = _AGEN_mainform.Data_Table_crossings.Rows.Count - 1; i >= 0; --i)
                                        {
                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["X"] != DBNull.Value && _AGEN_mainform.Data_Table_crossings.Rows[i]["Y"] != DBNull.Value)
                                            {
                                                double x1 = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i]["X"]);
                                                double y1 = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i]["Y"]);

                                                Point3d pt_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);


                                                double dist = poly2d.GetClosestPointTo(pt_on_poly1, Vector3d.ZAxis, false).DistanceTo(new Point3d(x1, y1, poly2d.Elevation));
                                                if (dist > 1)
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i].Delete();
                                                }
                                                else
                                                {
                                                    double sta1 = poly2d.GetDistAtPoint(pt_on_poly1);

                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta] != DBNull.Value)
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i]["EqSta"] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                        }
                                                        _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta] = DBNull.Value;
                                                    }

                                                    else
                                                    {
                                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                                        {
                                                            double param1 = poly2d.GetParameterAtPoint(pt_on_poly1);
                                                            sta1 = poly3d.GetDistanceAtParameter(param1);
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta] = Math.Round(sta1, _AGEN_mainform.round1);
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i]["EqSta"] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);

                                                            _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta] = DBNull.Value;
                                                        }
                                                    }

                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["X"] = pt_on_poly1.X;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Y"] = pt_on_poly1.Y;
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i]["Z"] = 0;
                                                }

                                            }
                                            else
                                            {

                                                Trans1.Commit();
                                                _AGEN_mainform.tpage_processing.Hide();
                                                Populate_datagridview_with_equation_data();

                                                set_enable_true();

                                                MessageBox.Show("crossing file contains a crossing that is not specified by x,y position\r\ncrossing file is not modified\r\noperation aborted");
                                                return;
                                            }

                                        }
                                    }
                                }

                                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                                {
                                    _AGEN_mainform.tpage_crossing_scan.Populate_crossing_file(fisier_cs);
                                }
                                #endregion


                                #region property xlxs   
                                if (_AGEN_mainform.Data_Table_property != null)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {
                                        if (_AGEN_mainform.Data_Table_property.Rows.Count > 0)
                                        {
                                            for (int i = _AGEN_mainform.Data_Table_property.Rows.Count - 1; i >= 0; --i)
                                            {
                                                if (_AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"] != DBNull.Value &&
                                                    _AGEN_mainform.Data_Table_property.Rows[i]["X_End"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["Y_End"] != DBNull.Value)
                                                {
                                                    double x1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"]);
                                                    double y1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"]);

                                                    double x2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["X_End"]);
                                                    double y2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["Y_End"]);

                                                    Point3d pt_on_poly1 = poly2d.GetClosestPointTo(new Point3d(x1, y1, poly2d.Elevation), Vector3d.ZAxis, false);
                                                    Point3d pt_on_poly2 = poly2d.GetClosestPointTo(new Point3d(x2, y2, poly2d.Elevation), Vector3d.ZAxis, false);

                                                    double dist1 = poly2d.GetClosestPointTo(pt_on_poly1, Vector3d.ZAxis, false).DistanceTo(new Point3d(x1, y1, poly2d.Elevation));
                                                    double dist2 = poly2d.GetClosestPointTo(pt_on_poly2, Vector3d.ZAxis, false).DistanceTo(new Point3d(x1, y1, poly2d.Elevation));
                                                    if (dist1 > 1)
                                                    {
                                                        _AGEN_mainform.Data_Table_property.Rows[i].Delete();
                                                    }
                                                    else
                                                    {
                                                        double sta1 = poly2d.GetDistAtPoint(pt_on_poly1);
                                                        double sta2 = poly2d.GetDistAtPoint(pt_on_poly2);
                                                        if (_AGEN_mainform.Project_type == "2D")
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i]["2DStaBeg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["2DStaEnd"] != DBNull.Value)
                                                            {

                                                                _AGEN_mainform.Data_Table_property.Rows[i]["2DStaBeg"] = Math.Round(sta1, _AGEN_mainform.round1);
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["EqStaBeg"] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);


                                                                _AGEN_mainform.Data_Table_property.Rows[i]["2DStaEnd"] = Math.Round(sta2, _AGEN_mainform.round1);
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["EqStaEnd"] = Math.Round(Functions.Station_equation_ofV2(sta2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);

                                                                _AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"] = DBNull.Value;
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"] = DBNull.Value;
                                                            }

                                                        }

                                                        else
                                                        {
                                                            if (_AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"] != DBNull.Value)
                                                            {
                                                                double param1 = poly2d.GetParameterAtPoint(pt_on_poly1);
                                                                double param2 = poly2d.GetParameterAtPoint(pt_on_poly2);
                                                                sta1 = poly3d.GetDistanceAtParameter(param1);
                                                                sta2 = poly3d.GetDistanceAtParameter(param2);

                                                                _AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"] = Math.Round(sta1, _AGEN_mainform.round1);
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["EqStaBeg"] = Math.Round(Functions.Station_equation_ofV2(sta1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);

                                                                _AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"] = Math.Round(sta2, _AGEN_mainform.round1);
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["EqStaEnd"] = Math.Round(Functions.Station_equation_ofV2(sta2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);


                                                                _AGEN_mainform.Data_Table_property.Rows[i]["2DStaBeg"] = DBNull.Value;
                                                                _AGEN_mainform.Data_Table_property.Rows[i]["2DStaEnd"] = DBNull.Value;
                                                            }
                                                        }
                                                        _AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"] = pt_on_poly1.X;
                                                        _AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"] = pt_on_poly1.Y;
                                                        _AGEN_mainform.Data_Table_property.Rows[i]["X_End"] = pt_on_poly2.X;
                                                        _AGEN_mainform.Data_Table_property.Rows[i]["Y_End"] = pt_on_poly2.Y;
                                                    }
                                                }
                                                else
                                                {
                                                    Trans1.Commit();
                                                    _AGEN_mainform.tpage_processing.Hide();
                                                    Populate_datagridview_with_equation_data();
                                                    set_enable_true();
                                                    MessageBox.Show("property file contains an entry that is not specified by x,y position\r\noperation aborted\r\nnothing was updated");
                                                    return;
                                                }

                                            }
                                        }
                                    }
                                }

                                if (_AGEN_mainform.Data_Table_property.Rows.Count > 0)
                                {
                                    _AGEN_mainform.tpage_owner_scan.Populate_property_file(fisier_prop);
                                }
                                #endregion

                                if (lista1.Count > 0)
                                {
                                    for (int i = 0; i < lista1.Count; ++i)
                                    {
                                        Entity ent1 = Trans1.GetObject(lista1[i].ObjectId, OpenMode.ForWrite) as Entity;
                                        if (ent1 != null)
                                        {
                                            ent1.Erase();
                                        }
                                    }
                                }
                                System.Data.DataColumn column1 = _AGEN_mainform.dt_station_equation.Columns["measured"];
                                _AGEN_mainform.dt_station_equation.Columns.Remove(column1);
                                if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();

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

            _AGEN_mainform.tpage_processing.Hide();
            Populate_datagridview_with_equation_data();

            set_enable_true();
        }



        public void Populate_datagridview_with_equation_data()
        {
            if (_AGEN_mainform.COUNTRY == "USA" && _AGEN_mainform.dt_station_equation != null)
            {
                if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                {
                    dt_display = new System.Data.DataTable();
                    dt_display.Columns.Add("No", typeof(int));
                    dt_display.Columns.Add("Back", typeof(string));
                    dt_display.Columns.Add("Ahead", typeof(string));
                    dt_display.Columns.Add("Show in plan", typeof(bool));
                    for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                    {

                        dt_display.Rows.Add();
                        double bs = -1.23456789;

                        dt_display.Rows[i]["No"] = i;
                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"] != DBNull.Value)
                        {
                            bs = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Back"]);

                            dt_display.Rows[i]["Back"] = Functions.Get_chainage_from_double(bs, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                        }

                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"] != DBNull.Value)
                        {
                            double sa = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Station Ahead"]);

                            dt_display.Rows[i]["Ahead"] = Functions.Get_chainage_from_double(sa, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                            if (bs == 0 && sa != 0) textBox_start_station_CL.Text = Functions.Get_chainage_from_double(sa, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                        }


                        bool show1 = true;
                        if (_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(_AGEN_mainform.dt_station_equation.Rows[i]["Show in plan"]);
                            if (val1.ToUpper() == "NO" || val1.ToUpper() == "FALSE")
                            {
                                show1 = false;
                            }




                        }
                        dt_display.Rows[i]["Show in plan"] = show1;

                    }



                }
            }

            display_station_equations(dt_display);

        }


        private void display_station_equations(System.Data.DataTable display_dt)
        {
            dataGridView_sta_eq.DataSource = display_dt;
            dataGridView_sta_eq.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_sta_eq.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sta_eq.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_sta_eq.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sta_eq.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_sta_eq.EnableHeadersVisualStyles = false;
        }

        private void button_redefine_begin_station_Click(object sender, EventArgs e)
        {
            Ag = this.MdiParent as _AGEN_mainform;

            

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.cl_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.cl_excel_name + " file");
                return;
            }

            set_enable_false();
            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }

            string Start_sta_cl = textBox_start_station_CL.Text;

            if (Functions.IsNumeric(Start_sta_cl.Replace("+", "")) == true)
            {
                double start1 = Convert.ToDouble(Start_sta_cl.Replace("+", ""));

                if (System.IO.Directory.Exists(ProjF) == true)
                {
                    string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        try
                        {


                            using (DocumentLock lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {

                                    Polyline3d poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                    if (poly3d != null)
                                    {

                                        bool adauga_new_entry = true;


                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows[0]["Station Back"] != DBNull.Value)
                                            {
                                                double bs = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[0]["Station Back"]);

                                                if (bs == 0)
                                                {
                                                    double sa_old = 0;
                                                    if (_AGEN_mainform.dt_station_equation.Rows[0]["Station Ahead"] != DBNull.Value)
                                                    {
                                                        sa_old = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[0]["Station Ahead"]);

                                                    }

                                                    if (start1 != 0)
                                                    {
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Station Ahead"] = start1;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute Start X"] = poly3d.StartPoint.X;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute Start Y"] = poly3d.StartPoint.Y;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute Start Z"] = poly3d.StartPoint.Z;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute End X"] = poly3d.StartPoint.X;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute End Y"] = poly3d.StartPoint.Y;
                                                        _AGEN_mainform.dt_station_equation.Rows[0]["Reroute End Z"] = poly3d.StartPoint.Z;


                                                        // here!!!!
                                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 1)
                                                        {

                                                            if (_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] != DBNull.Value)
                                                            {
                                                                double sb = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"]);
                                                                _AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] = sb + start1 - sa_old;
                                                            }


                                                        }

                                                    }
                                                    else
                                                    {
                                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 1)
                                                        {

                                                            if (_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] != DBNull.Value)
                                                            {
                                                                double sb = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"]);
                                                                _AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] = sb - sa_old;
                                                            }


                                                        }

                                                        _AGEN_mainform.dt_station_equation.Rows[0].Delete();
                                                    }
                                                    adauga_new_entry = false;
                                                }
                                            }
                                        }

                                        if (adauga_new_entry == true)
                                        {
                                            if (start1 != 0)
                                            {
                                                _AGEN_mainform.dt_station_equation.Rows.Add();
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Station Back"] = 0;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Station Ahead"] = start1;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start X"] = poly3d.StartPoint.X;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start Y"] = poly3d.StartPoint.Y;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute Start Z"] = poly3d.StartPoint.Z;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End X"] = poly3d.StartPoint.X;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End Y"] = poly3d.StartPoint.Y;
                                                _AGEN_mainform.dt_station_equation.Rows[_AGEN_mainform.dt_station_equation.Rows.Count - 1]["Reroute End Z"] = poly3d.StartPoint.Z;

                                                _AGEN_mainform.dt_station_equation = Functions.Sort_data_table(_AGEN_mainform.dt_station_equation, "Station Back");

                                                if (_AGEN_mainform.dt_station_equation.Rows.Count > 1)
                                                {
                                                    if (_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] != DBNull.Value)
                                                    {
                                                        double sb = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[1]["Station Back"]);
                                                        _AGEN_mainform.dt_station_equation.Rows[1]["Station Back"] = sb + start1;
                                                    }
                                                }
                                            }
                                        }

                                        Functions.create_backup(fisier_cl);
                                        _AGEN_mainform.tpage_setup.Add_to_centerline_file_station_equations(fisier_cl, _AGEN_mainform.dt_station_equation);


                                        string fisier_cs = ProjF + _AGEN_mainform.crossing_excel_name;

                                        if (System.IO.File.Exists(fisier_cs) == true)
                                        {

                                            if (System.IO.File.Exists(fisier_cs) == true)
                                            {
                                                Functions.create_backup(fisier_cs);
                                                _AGEN_mainform.Data_Table_crossings = _AGEN_mainform.tpage_crossing_scan.Load_existing_crossing(fisier_cs);
                                            }
                                            else
                                            {
                                                _AGEN_mainform.Data_Table_crossings = Functions.Creaza_crossing_datatable_structure();
                                            }

                                            if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                                            {
                                                string Col_EqSta = "EqSta";

                                                for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                                {
                                                    _AGEN_mainform.Data_Table_crossings.Rows[i][Col_EqSta] = DBNull.Value;

                                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta] != DBNull.Value)
                                                    {
                                                        double sta2d = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta]);

                                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i][Col_EqSta] = Math.Round(Functions.Station_equation_of(sta2d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                        }
                                                    }

                                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                                    {
                                                        double sta3d = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta]);

                                                        if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                        {
                                                            _AGEN_mainform.Data_Table_crossings.Rows[i][Col_EqSta] = Math.Round(Functions.Station_equation_of(sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.round1);
                                                        }
                                                    }
                                                }

                                                _AGEN_mainform.tpage_crossing_scan.Populate_crossing_file(fisier_cs);
                                            }
                                        }

                                        poly3d.Erase();
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
            }
            Populate_datagridview_with_equation_data();
            set_enable_true();

        }

        public void set_textBox_start_station_CL(string continut)
        {
            textBox_start_station_CL.Text = continut;
        }

        private void Button_Station_eq_pick_sta_Click(object sender, EventArgs e)
        {

            _AGEN_mainform.tpage_processing.Show();
            Ag = this.MdiParent as _AGEN_mainform;
            Ag.WindowState = FormWindowState.Minimized;


            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_rez.MessageForAdding = "\nSelect new centerline:";
                    Prompt_rez.SingleOnly = true;
                    Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                    if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        set_enable_true();
                        Ag.WindowState = FormWindowState.Normal;
                        _AGEN_mainform.tpage_processing.Hide();
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        Editor1.SetImpliedSelection(Empty_array);
                        return;
                    }

                    System.Data.DataTable Data_table_station_equation = Functions.Creaza_station_equation_datatable_structure();


                    Polyline Poly2D = null;


                    string layer_rstart = "Agen Reroute Start";
                    string layer_bsta = "Agen Back Station";
                    string layer_asta = "Agen Ahead Station";

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Poly2D = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForRead) as Polyline;


                        if (Poly2D == null)
                        {
                            set_enable_true();
                            Ag.WindowState = FormWindowState.Normal;
                            _AGEN_mainform.tpage_processing.Hide();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            return;
                        }



                        Functions.Creaza_layer(layer_rstart, 2, false);
                        Functions.Creaza_layer(layer_bsta, 5, false);
                        Functions.Creaza_layer(layer_asta, 1, false);

                        Trans1.Commit();
                    }

                label_repeat:

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the reroute start point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the reroute end point");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the a Back station point");
                        PP3.AllowNone = false;
                        Point_res3 = Editor1.GetPoint(PP3);

                        if (Point_res3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_double = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify a Back station:");
                        Prompt_double.AllowNegative = false;
                        Prompt_double.AllowZero = true;
                        Prompt_double.AllowNone = true;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_double3 = ThisDrawing.Editor.GetDouble(Prompt_double);
                        if (Rezultat_double3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP4;
                        PP4 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the an Ahead station point");
                        PP4.AllowNone = false;

                        Point_res4 = Editor1.GetPoint(PP4);

                        if (Point_res4.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_double1 = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify an Ahead station:");
                        Prompt_double1.AllowNegative = false;
                        Prompt_double1.AllowZero = true;
                        Prompt_double1.AllowNone = true;
                        Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_double4 = ThisDrawing.Editor.GetDouble(Prompt_double1);
                        if (Rezultat_double4.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            goto label_delete;
                        }


                        double sta3 = Rezultat_double3.Value;
                        double sta4 = Rezultat_double4.Value;

                        Point3d ptnew1 = Poly2D.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);
                        Point3d ptnew2 = Poly2D.GetClosestPointTo(Point_res2.Value, Vector3d.ZAxis, false);
                        Point3d ptnew3 = Poly2D.GetClosestPointTo(Point_res3.Value, Vector3d.ZAxis, false);
                        Point3d ptnew4 = Poly2D.GetClosestPointTo(Point_res4.Value, Vector3d.ZAxis, false);




                        double d1 = Poly2D.GetDistAtPoint(ptnew1);
                        double d2 = Poly2D.GetDistAtPoint(ptnew2);
                        double d3 = Poly2D.GetDistAtPoint(ptnew3);
                        double d4 = Poly2D.GetDistAtPoint(ptnew4);

                        if (d1 > d2)
                        {
                            double t = d1;
                            d1 = d2;
                            d2 = t;
                        }


                        double r1 = Math.Round(sta3 + (d1 - d3), 3);


                        double Back11 = Math.Round(sta3 + (d2 - d3), 3);


                        double Ahead11 = Math.Round(sta4 + (d2 - d4), 3);

                        Data_table_station_equation.Rows.Add();
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute Start X"] = ptnew1.X;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute Start Y"] = ptnew1.Y;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute Start Z"] = 0;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute End X"] = ptnew2.X;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute End Y"] = ptnew2.Y;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Reroute End Z"] = 0;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Station Back"] = Back11;
                        Data_table_station_equation.Rows[Data_table_station_equation.Rows.Count - 1]["Station Ahead"] = Ahead11;


                        MLeader Ml1 = Functions.creaza_mleader(ptnew1, Functions.Get_chainage_from_double(r1, _AGEN_mainform.units_of_measurement, 0), 50, 50, 50, 20, 20, 2.5);
                        Ml1.Layer = layer_rstart;
                        Ml1.ColorIndex = 256;

                        MLeader Ml2 = Functions.creaza_mleader(ptnew2, Functions.Get_chainage_from_double(Back11, _AGEN_mainform.units_of_measurement, 0), 50, 50, 50, 20, 20, 2.5);
                        Ml2.Layer = layer_bsta;
                        Ml2.ColorIndex = 256;

                        MLeader Ml3 = Functions.creaza_mleader(ptnew2, Functions.Get_chainage_from_double(Ahead11, _AGEN_mainform.units_of_measurement, 0), 50, 50, -50, 20, 20, 2.5);
                        Ml3.Layer = layer_asta;
                        Ml3.ColorIndex = 256;

                        Trans1.Commit();
                        goto label_repeat;
                    }

                label_delete:
                    Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Data_table_station_equation);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            MessageBox.Show("done");

            _AGEN_mainform.tpage_processing.Hide();

            set_enable_true();




        }

        private void button_add_prefix_r2_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the mp blocks:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as BlockReference;
                            if (block1 != null)
                            {
                                if (block1.AttributeCollection.Count > 0)
                                {
                                    foreach (ObjectId id1 in block1.AttributeCollection)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForWrite) as AttributeReference;
                                        if (atr1 != null)
                                        {
                                            if (atr1.Tag.ToLower() == "kp" || atr1.Tag.ToLower() == "mp")
                                            {
                                                string val1 = atr1.TextString;
                                                if (val1.Contains("R") == true)
                                                {
                                                    int index1 = val1.IndexOf("R");
                                                    val1 = val1.Substring(0, index1);
                                                }


                                                atr1.TextString = atr1.TextString + textBoxR.Text + textBox_no.Text;
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");


        }

        private void Label_station_equations_Click(object sender, EventArgs e)
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

        private void button_transf_to_excel_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (dt_display != null && dt_display.Rows.Count > 0)
                {
                    if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt_display.Rows.Count; ++i)
                        {
                            if (dt_display.Rows[i][1] != DBNull.Value && dt_display.Rows[i][2] != DBNull.Value && dt_display.Rows[i][3] != DBNull.Value)
                            {
                                string back1_s = Convert.ToString(dt_display.Rows[i][1]);
                                string ahead1_s = Convert.ToString(dt_display.Rows[i][2]);
                                bool show1_b = Convert.ToBoolean(dt_display.Rows[i][3]);

                                string show1 = "YES";
                                if (show1_b == false) show1 = "NO";

                                for (int j = 0; j < _AGEN_mainform.dt_station_equation.Rows.Count; ++j)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows[j][5] != DBNull.Value &&
                                        _AGEN_mainform.dt_station_equation.Rows[j][6] != DBNull.Value)
                                    {
                                        double back2 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[j][5]);
                                        double ahead2 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[j][6]);
                                        string back2_s = Functions.Get_chainage_from_double(back2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                        string ahead2_s = Functions.Get_chainage_from_double(ahead2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                        if (ahead2_s == ahead1_s && back2_s == back1_s)
                                        {
                                            _AGEN_mainform.dt_station_equation.Rows[j][11] = show1;
                                            j = _AGEN_mainform.dt_station_equation.Rows.Count;
                                        }
                                    }
                                }
                            }
                        }

                        string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }
                        string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                        if (System.IO.File.Exists(fisier_cl) == true)
                        {
                            Populate_station_eq_file(fisier_cl, false);
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





        public void Populate_station_eq_file(string File1, bool delete_steq)
        {
            try
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

                Excel1.Visible = _AGEN_mainform.ExcelVisible;
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
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_centerline, _AGEN_mainform.Start_row_CL, "General");
                    Functions.Create_header_centerline_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1, false, _AGEN_mainform.version);

                    if (delete_steq == true)
                    {
                        _AGEN_mainform.dt_station_equation = Functions.Creaza_station_equation_datatable_structure();
                        if (Workbook1.Worksheets.Count > 1)
                        {
                            Workbook1.Worksheets[2].Columns["A:XX"].Delete();
                        }
                    }
                    else
                    {
                        if (_AGEN_mainform.dt_station_equation != null)
                        {
                            if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                            {

                                if (Workbook1.Worksheets.Count == 1)
                                {
                                    Microsoft.Office.Interop.Excel.Worksheet W3 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(W3);
                                }

                                Microsoft.Office.Interop.Excel.Worksheet W2 = Workbook1.Worksheets[2];
                                W2.Name = "St_eq";

                                try
                                {
                                    Functions.Transfer_to_worksheet_Data_table(W2, _AGEN_mainform.dt_station_equation, _AGEN_mainform.Start_row_station_equation, "General");
                                    Functions.Create_header_station_eq(W2, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);
                                }
                                finally
                                {
                                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);

                                }
                            }
                        }

                    }

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
    }
}
