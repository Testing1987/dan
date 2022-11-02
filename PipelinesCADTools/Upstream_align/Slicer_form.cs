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

namespace Alignment_mdi
{
    public partial class slicer_form : Form
    {
        bool Freeze_operations = false;
        System.Data.DataTable dtcl = null;
        System.Data.DataTable[] dt1 = null;
        System.Data.DataTable dt2 = null;
        double ang_precision;
        int no_radial_lines;
        double vert_precision = 0.1;
        int no_vert_lines = 401;


        #region mouse move exit and minimize
        bool clickdragdown;
        Point lastLocation;
        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown == true)
            {
                this.Location = new Point(
                  (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void clickmove_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = false;
        }
        private void button_Exit_Click(object sender, EventArgs e)
        {


            this.Close();
        }
        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        #endregion
        public slicer_form()
        {
            InitializeComponent();

        }



        private void button_select_centerline_Click(object sender, EventArgs e)
        {
            dtcl = null;
            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;



                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a 3D polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                            Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                            if (Rezultat_centerline.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                Editor1.SetImpliedSelection(Empty_array);
                                return;
                            }

                            Polyline3d poly3d = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;
                            if (poly3d != null)
                            {
                                dtcl = new System.Data.DataTable();
                                dtcl.Columns.Add("index", typeof(int));
                                dtcl.Columns.Add("x", typeof(double));
                                dtcl.Columns.Add("y", typeof(double));
                                dtcl.Columns.Add("z", typeof(double));
                                dtcl.Columns.Add("sta", typeof(double));
                                for (int i = 0; i <= poly3d.EndParam; ++i)
                                {
                                    dtcl.Rows.Add();
                                    dtcl.Rows[dtcl.Rows.Count - 1][0] = i;
                                    dtcl.Rows[dtcl.Rows.Count - 1][1] = poly3d.GetPointAtParameter(i).X;
                                    dtcl.Rows[dtcl.Rows.Count - 1][2] = poly3d.GetPointAtParameter(i).Y;
                                    dtcl.Rows[dtcl.Rows.Count - 1][3] = poly3d.GetPointAtParameter(i).Z;
                                    dtcl.Rows[dtcl.Rows.Count - 1][4] = poly3d.GetDistanceAtParameter(i);
                                }
                                button_cl_loaded.Visible = true;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    button_cl_loaded.Visible = false;
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }

        }

        private void button_select_slices_Click(object sender, EventArgs e)
        {
            string start_string = textBox_start_station.Text;
            if (Functions.IsNumeric(start_string.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified properly");
                return;
            }

            string exclusion = textBox_z_exclude.Text;
            if (Functions.IsNumeric(exclusion) == false)
            {
                MessageBox.Show("please enter a numeric value");
                return;
            }

            double deltaZ = Math.Abs(Convert.ToDouble(exclusion));

            double start_sta = Convert.ToDouble(start_string.Replace("+", ""));
            button_slices_loaded.Visible = false;
            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                if (checkBox_append_slices.Checked == false)
                {
                    dt1 = null;
                    dt2 = null;
                }
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Editor1.SetImpliedSelection(Empty_array);
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    if (Functions.IsNumeric(textBox_scanning_precision.Text) == false)
                    {
                        Editor1.WriteMessage("\nCommand:");
                        MessageBox.Show("not numeric precision");
                        return;
                    }
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect slices:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                Freeze_operations = false;
                                return;
                            }


                            Polyline3d Poly3D = new Polyline3d();
                            Poly3D.SetDatabaseDefaults();
                            Poly3D.Layer = "0";
                            BTrecord.AppendEntity(Poly3D);
                            Trans1.AddNewlyCreatedDBObject(Poly3D, true);
                            Build_3d_poly_from_datatable(Trans1, Poly3D, dtcl);


                            ang_precision = Math.Abs(Convert.ToDouble(textBox_scanning_precision.Text));
                            no_radial_lines = Convert.ToInt32(180 / ang_precision) + 1;


                            if (dt1 == null)
                            {
                                dt1 = new System.Data.DataTable[no_radial_lines];

                                for (int i = 0; i < no_radial_lines; ++i)
                                {
                                    dt1[i] = new System.Data.DataTable();
                                    dt1[i].Columns.Add("x", typeof(double));
                                    dt1[i].Columns.Add("y", typeof(double));
                                    dt1[i].Columns.Add("z", typeof(double));
                                    dt1[i].Columns.Add("sta", typeof(double));
                                }
                            }


                            if (dt2 == null)
                            {

                                dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("x", typeof(double));
                                dt2.Columns.Add("y", typeof(double));
                                dt2.Columns.Add("z", typeof(double));
                                dt2.Columns.Add("param", typeof(double));
                            }

                            int start_for_param_dt2 = dt2.Rows.Count;


                            List<double> lista_z = new List<double>();



                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline3d Poly3d_slice = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline3d;
                                if (Poly3d_slice != null)
                                {
                                    double Xmax = Poly3d_slice.GetPointAtParameter(0).X;
                                    double Ymax = Poly3d_slice.GetPointAtParameter(0).Y;
                                    double Zmin = Poly3d_slice.GetPointAtParameter(0).Z;

                                    for (int j = 0; j <= Poly3d_slice.EndParam; ++j)
                                    {
                                        double X = Math.Abs(Poly3d_slice.GetPointAtParameter(j).X);
                                        if (X > Xmax) Xmax = X;

                                        double Y = Math.Abs(Poly3d_slice.GetPointAtParameter(j).Y);
                                        if (Y > Ymax) Ymax = Y;

                                        double Z = Poly3d_slice.GetPointAtParameter(j).Z;
                                        if (Z < Zmin) Zmin = Z;
                                    }

                                    bool proceseraza = true;
                                    if (lista_z.Count > 0)

                                    {
                                        for (int r = 0; r < lista_z.Count; ++r)
                                        {
                                            if (Math.Abs(lista_z[r] - Zmin) <= deltaZ)
                                            {
                                                proceseraza = false;
                                                r = lista_z.Count;
                                            }
                                        }
                                    }
                                    lista_z.Add(Zmin);
                                    #region deltaz
                                    if (proceseraza == true)
                                    {
                                        double angle = 0;
                                        Point3d point0 = new Point3d(0, 0, Zmin);

                                        Polyline poly2d = new Polyline();
                                        poly2d = append_vertices_from_3Dpoly(poly2d, Poly3d_slice, Zmin);


                                        Polyline l1 = new Polyline();
                                        l1.Elevation = Zmin;
                                        l1.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
                                        l1.AddVertexAt(1, new Point2d(Xmax + Ymax, 0), 0, 0, 0);


                                        for (int m = 0; m < no_radial_lines; ++m)
                                        {
                                            if (m == 0)
                                            {
                                                for (int p = 0; p < no_vert_lines; ++p)
                                                {
                                                    if (-20 + p * vert_precision <= -3 | -20 + p * vert_precision >= 3)
                                                    {
                                                        Polyline l2 = new Polyline();
                                                        l2.Elevation = Zmin;
                                                        l2.AddVertexAt(0, new Point2d(-20 + p * vert_precision, 0), 0, 0, 0);
                                                        l2.AddVertexAt(1, new Point2d(-20 + p * vert_precision, -100), 0, 0, 0);
                                                        Point3dCollection colint2 = Functions.Intersect_on_both_operands(poly2d, l2);

                                                        if (colint2.Count > 0)
                                                        {
                                                            for (int q = 0; q < colint2.Count; ++q)
                                                            {
                                                                double param2 = poly2d.GetParameterAtPoint(colint2[q]);
                                                                dt2.Rows.Add();
                                                                dt2.Rows[dt2.Rows.Count - 1][3] = param2;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            if (m > 0)
                                            {
                                                poly2d = new Polyline();
                                                poly2d = append_vertices_from_3Dpoly(poly2d, Poly3d_slice, Zmin);
                                                l1.TransformBy(Matrix3d.Rotation(ang_precision * Math.PI / 180, Vector3d.ZAxis, point0));
                                                angle = angle + ang_precision;
                                            }
                                            Point3dCollection colint = Functions.Intersect_on_both_operands(poly2d, l1);

                                            if (colint.Count == 0)
                                            {
                                                MessageBox.Show("the slice from " + Poly3d_slice.Layer + " is not intersecting the " + angle + " line");
                                                BTrecord.AppendEntity(l1);
                                                Trans1.AddNewlyCreatedDBObject(l1, true);

                                                Polyline pp_err = new Polyline();
                                                pp_err = append_vertices_from_3Dpoly(pp_err, Poly3d_slice, Zmin);
                                                pp_err.ColorIndex = 3;
                                                BTrecord.AppendEntity(pp_err);
                                                Trans1.AddNewlyCreatedDBObject(pp_err, true);

                                                Trans1.Commit();

                                                Editor1.SetImpliedSelection(Empty_array);
                                                Editor1.WriteMessage("\nCommand:");
                                                Freeze_operations = false;
                                                return;
                                            }

                                            Point3d point1 = colint[0];
                                            double dp = 0;
                                            if (point0.DistanceTo(point1) > 3)
                                            {
                                                dp = point0.DistanceTo(point1);
                                            }
                                            if (colint.Count > 1)
                                            {
                                                for (int k = 1; k < colint.Count; ++k)
                                                {
                                                    double d1 = point0.DistanceTo(colint[k]);
                                                    if (d1 > 3)
                                                    {
                                                        if (dp > d1 || dp == 0)
                                                        {
                                                            point1 = colint[k];
                                                            dp = d1;
                                                        }
                                                    }

                                                }
                                            }

                                            if (dp > 0)
                                            {
                                                double param1 = poly2d.GetParameterAtPoint(point1);

                                                align_poly_to_3d_poly(poly2d, Poly3D, 1, start_sta);
                                                Point3d pt_transfered = poly2d.GetPointAtParameter(param1);

                                                dt1[m].Rows.Add();
                                                dt1[m].Rows[dt1[m].Rows.Count - 1][0] = pt_transfered.X;
                                                dt1[m].Rows[dt1[m].Rows.Count - 1][1] = pt_transfered.Y;
                                                dt1[m].Rows[dt1[m].Rows.Count - 1][2] = pt_transfered.Z;
                                                dt1[m].Rows[dt1[m].Rows.Count - 1][3] = Zmin;

                                                for (int p = start_for_param_dt2; p < dt2.Rows.Count; ++p)
                                                {
                                                    double param2 = Convert.ToDouble(dt2.Rows[p][3]);
                                                    dt2.Rows[p][0] = poly2d.GetPointAtParameter(param2).X;
                                                    dt2.Rows[p][1] = poly2d.GetPointAtParameter(param2).Y;
                                                    dt2.Rows[p][2] = poly2d.GetPointAtParameter(param2).Z;
                                                }

                                                start_for_param_dt2 = dt2.Rows.Count;

                                            }

                                            if (m > 0)
                                            {
                                                poly2d.Erase();
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }

                            Poly3D.Erase();
                            Trans1.Commit();
                            button_slices_loaded.Visible = true;



                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    button_slices_loaded.Visible = false;
                }

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                MessageBox.Show("done");
                Freeze_operations = false;
            }
        }
        public Polyline append_vertices_from_3Dpoly(Polyline Poly2D, Polyline3d Poly3D, double Zmin)
        {

            Poly2D.Elevation = Zmin;

            if (Poly3D.Length > 0)
            {

                double last_param = Poly3D.EndParam;

                for (int i = 0; i <= last_param; ++i)
                {
                    Poly2D.AddVertexAt(i, new Point2d(Poly3D.GetPointAtParameter(i).X, Poly3D.GetPointAtParameter(i).Y), 0, 0, 0);

                }
            }
            return Poly2D;

        }

        public void Build_3d_poly_from_datatable(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Polyline3d Poly3D, System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    for (int i = 0; i < dt1.Rows.Count; ++i)
                    {
                        double x = Convert.ToDouble(dt1.Rows[i]["x"]);
                        double y = Convert.ToDouble(dt1.Rows[i]["y"]);
                        double z = Convert.ToDouble(dt1.Rows[i]["z"]);

                        PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                        Poly3D.AppendVertex(Vertex_new);
                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                    }
                }
            }

        }

        public void align_poly_to_3d_poly(Polyline poly2d, Polyline3d cl, int colorindex, double start_sta)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                if (cl != null)
                {
                    if (poly2d != null)
                    {
                        double Sta1 = 1000 * poly2d.GetPointAtParameter(0).Z - start_sta;
                        if (cl.Length < Sta1)
                        {
                            MessageBox.Show("the cl length is smaller than station you specified\r\noperation aborted");
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        Point3d pt1 = cl.GetPointAtDist(Sta1);
                        double param1 = cl.GetParameterAtDistance(Sta1);
                        double Low = Math.Floor(param1);
                        double high = Math.Ceiling(param1);

                        if (Low == high) Low = Low - 1;
                        if (Low <= 0)
                        {
                            Low = 0;
                            high = 1;
                        }

                        if (high == cl.EndParam) Low = cl.EndParam - 1;



                        double x1 = cl.GetPointAtParameter(Low).X;
                        double y1 = cl.GetPointAtParameter(Low).Y;
                        double z1 = cl.GetPointAtParameter(Low).Z;

                        double x2 = cl.GetPointAtParameter(high).X;
                        double y2 = cl.GetPointAtParameter(high).Y;
                        double z2 = cl.GetPointAtParameter(high).Z;



                        Point3d pt0 = new Point3d(0, 0, poly2d.GetPointAtParameter(0).Z);



                        poly2d.TransformBy(Matrix3d.Displacement(pt0.GetVectorTo(pt1)));

                        double roth = Functions.GET_Bearing_rad(x1, y1, x2, y2) - Math.PI / 2;
                        poly2d.TransformBy(Matrix3d.Rotation(roth, Vector3d.ZAxis, pt1));

                        Vector3d v1 = new Point3d(x1, y1, z1).GetVectorTo(new Point3d(x2, y2, z2));
                        Vector3d v2 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));

                        double angv = v1.GetAngleTo(v2);

                        Vector3d v11 = new Point3d(x1, y1, z1).GetVectorTo(new Point3d(x2, y2, z1));
                        v11 = v11.GetNormal();
                        Vector3d cp = -v1.CrossProduct(v11);

                        double xtra = 0;
                        if (Math.Round(z1, 4) > Math.Round(z2, 4))
                        {
                            xtra = Math.PI;
                        }
                        else if (Math.Round(z1, 4) == Math.Round(z2, 4))
                        {
                            Autodesk.AutoCAD.DatabaseServices.Line l1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(x1, y1, z1), new Point3d(x2, y2, z2));
                            l1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, l1.StartPoint));
                            cp = -l1.StartPoint.GetVectorTo(l1.EndPoint);

                        }

                        poly2d.TransformBy(Matrix3d.Rotation(xtra + angv + Math.PI / 2, cp, pt1));
                        poly2d.Layer = "0";
                        poly2d.ColorIndex = colorindex;
                        BTrecord.AppendEntity(poly2d);
                        Trans1.AddNewlyCreatedDBObject(poly2d, true);

                    }
                    Trans1.Commit();
                }
            }
        }

        private void button_generate_aligned_polylines_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    if (button_slices_loaded.Visible == false)
                    {

                        Editor1.WriteMessage("\nCommand:");
                        MessageBox.Show("slices not loaded");
                        return;
                    }

                    if (button_cl_loaded.Visible == false)
                    {

                        Editor1.WriteMessage("\nCommand:");
                        MessageBox.Show("centerline not loaded");
                        return;
                    }
                    Freeze_operations = true;

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline3d Poly3D = new Polyline3d();
                            Poly3D.SetDatabaseDefaults();
                            Poly3D.Layer = "0";
                            BTrecord.AppendEntity(Poly3D);
                            Trans1.AddNewlyCreatedDBObject(Poly3D, true);
                            Build_3d_poly_from_datatable(Trans1, Poly3D, dtcl);

                            Int16 ci = 1;
                            for (int k = 0; k < no_radial_lines; ++k)
                            {
                                dt1[k] = Functions.Sort_data_table(dt1[k], "sta");

                                string layername = Functions.Get_String_Rounded(k * ang_precision, 1);

                                Functions.Creaza_layer(layername, ci, true);

                                Polyline3d Poly3D_align = new Polyline3d();
                                Poly3D_align.SetDatabaseDefaults();
                                Poly3D_align.Layer = layername;
                                Poly3D_align.ColorIndex = 256;
                                BTrecord.AppendEntity(Poly3D_align);
                                Trans1.AddNewlyCreatedDBObject(Poly3D_align, true);
                                Build_3d_poly_from_datatable(Trans1, Poly3D_align, dt1[k]);

                                ci = Convert.ToInt16(ci + 1);
                                if (ci == 8) ci = 1;
                            }

                            Functions.Creaza_layer("_bottom pts", 1, true);

                            for (int p = 0; p < dt2.Rows.Count; ++p)
                            {
                                Point3d pt2 = new Point3d(Convert.ToDouble(dt2.Rows[p][0]), Convert.ToDouble(dt2.Rows[p][1]), Convert.ToDouble(dt2.Rows[p][2]));

                                DBPoint dbpt2 = new DBPoint(pt2);
                                dbpt2.Layer = "_bottom pts";
                                dbpt2.ColorIndex = 256;
                                BTrecord.AppendEntity(dbpt2);
                                Trans1.AddNewlyCreatedDBObject(dbpt2, true);

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
                Freeze_operations = false;
            }
        }

        private void button_place_stations_along_cl_Click(object sender, EventArgs e)
        {
            string start_string = textBox_start_station.Text;
            if (Functions.IsNumeric(start_string.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified properly");
                return;
            }

            double start_sta = Convert.ToDouble(start_string.Replace("+", ""));
            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Editor1.SetImpliedSelection(Empty_array);
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {

                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect slices:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                Freeze_operations = false;
                                return;
                            }


                            Polyline3d Poly3D = new Polyline3d();
                            Poly3D.SetDatabaseDefaults();
                            Poly3D.Layer = "0";
                            BTrecord.AppendEntity(Poly3D);
                            Trans1.AddNewlyCreatedDBObject(Poly3D, true);
                            Build_3d_poly_from_datatable(Trans1, Poly3D, dtcl);


                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline poly_slice_aligned = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                if (poly_slice_aligned != null)
                                {
                                    Point3d point_on_poly = Poly3D.GetClosestPointTo(poly_slice_aligned.GetPointAtParameter(0), Vector3d.ZAxis, false);
                                    double sta = start_sta + Poly3D.GetDistAtPoint(point_on_poly);

                                    string stationstring = Functions.Get_chainage_from_double(sta, "f", 0);
                                    MText Mt_sta = new MText();
                                    Mt_sta.Normal = poly_slice_aligned.Normal;

                                    Mt_sta.Contents = stationstring;
                                    Mt_sta.Layer = "0";
                                    Mt_sta.Attachment = AttachmentPoint.BottomCenter;
                                    Mt_sta.TextHeight = 0.2;
                                    Mt_sta.ColorIndex = 256;
                                    Mt_sta.Location = point_on_poly;
                                    BTrecord.AppendEntity(Mt_sta);
                                    Trans1.AddNewlyCreatedDBObject(Mt_sta, true);


                                }
                            }

                            Poly3D.Erase();
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
                Freeze_operations = false;
            }
        }

        private void button_loft_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            System.Data.DataTable dt1 = new System.Data.DataTable();
                            for (int i = 0; i <= 180; ++i)
                            {
                                string col_name = Alignment_mdi.Functions.Get_String_Rounded(i, 1);
                                dt1.Columns.Add(col_name, typeof(Point3d));
                            }
                            bool add_row = true;

                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline3d poly3 = Trans1.GetObject(id1, OpenMode.ForWrite) as Polyline3d;
                                if (poly3 != null)
                                {
                                    if (dt1.Columns.Contains(poly3.Layer) == true)
                                    {
                                        if (dt1.Rows.Count > 0) add_row = false;
                                        for (int i = 0; i <= poly3.EndParam; ++i)
                                        {
                                            if (add_row == true) dt1.Rows.Add();
                                            dt1.Rows[i][poly3.Layer] = poly3.GetPointAtParameter(i);
                                        }
                                    }
                                }
                            }
                            Alignment_mdi.Functions.Creaza_layer("_optimized_slices", 200, true);

                            LoftProfile[] lps = new LoftProfile[2];

                            LoftProfile lp1 = new LoftProfile();

                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {

                                Polyline3d new_slice = new Polyline3d();
                                new_slice.Layer = "_optimized_slices";
                                new_slice.ColorIndex = 256;
                                BTrecord.AppendEntity(new_slice);
                                Trans1.AddNewlyCreatedDBObject(new_slice, true);

                                new_slice.SetDatabaseDefaults();

                                for (int j = 0; j < dt1.Columns.Count; ++j)
                                {
                                    Point3d pt1 = (Point3d)dt1.Rows[i][j];
                                    PolylineVertex3d Vertex_new = new PolylineVertex3d(pt1);
                                    new_slice.AppendVertex(Vertex_new);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                }

                                if (i != 1424)
                                {
                                    if (i > 0)
                                    {
                                        LoftProfile lp2 = new LoftProfile(new_slice);
                                        lps[1] = lp2;
                                        try
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.Surface.CreateLoftedSurface(lps, null, null, new LoftOptions(), true);
                                        }
                                        catch (System.Exception ex)
                                        {
                                            MessageBox.Show(ex.Message + "\r\nExcel Row:" + Convert.ToString(i + 2));
                                        }

                                        lp1 = new LoftProfile(new_slice);
                                        lps[0] = lp1;


                                    }
                                    else
                                    {
                                        lp1 = new LoftProfile(new_slice);
                                        lps[0] = lp1;
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
                Freeze_operations = false;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the 3dpolylines:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                            //for (int i=0; i< Rezultat1.Value.Count;++i)

                            Polyline3d p3d1 = Trans1.GetObject(Rezultat1.Value[0].ObjectId, OpenMode.ForWrite) as Polyline3d;
                            Polyline3d p3d2 = Trans1.GetObject(Rezultat1.Value[1].ObjectId, OpenMode.ForWrite) as Polyline3d;
                            Polyline3d p3d3 = Trans1.GetObject(Rezultat1.Value[2].ObjectId, OpenMode.ForWrite) as Polyline3d;

                            LoftProfile lp1 = new LoftProfile(p3d1);
                            LoftProfile lp2 = new LoftProfile(p3d2);
                            LoftProfile lp3 = new LoftProfile(p3d3);

                            LoftProfile[] lps = new LoftProfile[3] { lp1, lp2, lp3 };



                            Autodesk.AutoCAD.DatabaseServices.Surface.CreateLoftedSurface(lps, null, null, new LoftOptions(), true);

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
                Freeze_operations = false;
            }

        }

        private void button_export_data_to_excel_Click(object sender, EventArgs e)
        {


            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                            System.Data.DataTable dt1 = new System.Data.DataTable();
                            for (int i = 0; i <= 180; ++i)
                            {
                                string col_name = Alignment_mdi.Functions.Get_String_Rounded(i, 1);
                                dt1.Columns.Add(col_name, typeof(Point3d));
                            }
                            bool add_row = true;

                            foreach (ObjectId id1 in BTrecord)
                            {
                                Polyline3d poly3 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline3d;
                                if (poly3 != null)
                                {
                                    if (dt1.Columns.Contains(poly3.Layer) == true)
                                    {
                                        if (dt1.Rows.Count > 0) add_row = false;
                                        for (int i = 0; i <= poly3.EndParam; ++i)
                                        {
                                            if (add_row == true) dt1.Rows.Add();
                                            dt1.Rows[i][poly3.Layer] = poly3.GetPointAtParameter(i);
                                        }
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

                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }

        }

    }
}
