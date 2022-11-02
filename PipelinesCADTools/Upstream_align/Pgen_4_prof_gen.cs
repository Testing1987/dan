using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
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
    public partial class Pgen_prof_gen3 : Form
    {
        List<string> scales;
        Pgen_mainform Pg = null;

        System.Data.DataTable dt_stream;
        System.Data.DataTable dt_top_of_bank_ne;
        System.Data.DataTable dt_top_of_bank_sw;
        System.Data.DataTable dt_cross;
        System.Data.DataTable dt_eq;
        Polyline refcl = null;

        double pipe_sta = -1.234;
        List<double> lista_lod_sta = null;

        public string layer_stationing = "Pgen_stationing";



        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(Button_Load_stream_cl);
            lista_butoane.Add(button_load_TOB_en);
            lista_butoane.Add(button_load_TOB_ws);

            lista_butoane.Add(Button_draw_prof_streamcl);
            lista_butoane.Add(button_LOD);
            lista_butoane.Add(button_pipe_int);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(Button_Load_stream_cl);
            lista_butoane.Add(button_load_TOB_en);
            lista_butoane.Add(button_load_TOB_ws);

            lista_butoane.Add(Button_draw_prof_streamcl);
            lista_butoane.Add(button_LOD);
            lista_butoane.Add(button_pipe_int);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Pgen_prof_gen3()
        {
            InitializeComponent();
        }

        private void pgen_label_page_load(object sender, EventArgs e)
        {
            scales = new List<string>();
            scales.Add("1:1");
            scales.Add("1:10");
            scales.Add("1:20");
            scales.Add("1:30");
            scales.Add("1:40");
            scales.Add("1:50");
            scales.Add("1:60");
            scales.Add("1:100");
            scales.Add("1:200");

            Pg = this.MdiParent as Pgen_mainform;
            if (Functions.is_dan_popescu() == true)
            {
                Pgen_mainform.ExcelVisible = true;
            }
            Combobox_scales.DataSource = scales;
            Combobox_scales.SelectedIndex = 1;
        }

        private System.Data.DataTable creaza_data_table(Polyline3d poly1, Polyline poly_ref, int prof_idx)
        {
            double max1 = -100000;
            double min1 = 100000;
            if (prof_idx == 1)
            {
                if (Functions.IsNumeric(textBox_el_bottom1.Text) == true && Functions.IsNumeric(textBox_el_top1.Text) == true)
                {
                    max1 = Convert.ToDouble(textBox_el_top1.Text);
                    min1 = Convert.ToDouble(textBox_el_bottom1.Text);
                }
            }

            if (prof_idx == 2)
            {
                if (Functions.IsNumeric(textBox_el_bottom2.Text) == true && Functions.IsNumeric(textBox_el_top2.Text) == true)
                {
                    max1 = Convert.ToDouble(textBox_el_top2.Text);
                    min1 = Convert.ToDouble(textBox_el_bottom2.Text);
                }
            }

            if (prof_idx == 3)
            {
                if (Functions.IsNumeric(textBox_el_bottom3.Text) == true && Functions.IsNumeric(textBox_el_top3.Text) == true)
                {
                    max1 = Convert.ToDouble(textBox_el_top3.Text);
                    min1 = Convert.ToDouble(textBox_el_bottom3.Text);
                }
            }

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("sta", typeof(double));
            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));
            dt1.Columns.Add("z", typeof(double));
            if (poly1 != null)
            {
                for (int i = 0; i <= poly1.EndParam; ++i)
                {
                    double x1 = poly1.GetPointAtParameter(i).X;
                    double y1 = poly1.GetPointAtParameter(i).Y;
                    if (poly_ref != null)
                    {
                        double z1 = poly1.GetPointAtParameter(i).Z;
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = poly_ref.GetDistAtPoint(poly_ref.GetClosestPointTo(new Point3d(x1, y1, poly_ref.Elevation), Vector3d.ZAxis, false));
                        dt1.Rows[dt1.Rows.Count - 1][1] = poly1.GetPointAtParameter(i).X;
                        dt1.Rows[dt1.Rows.Count - 1][2] = poly1.GetPointAtParameter(i).Y;
                        dt1.Rows[dt1.Rows.Count - 1][3] = z1;
                        if (z1 > max1 - 2) max1 = Functions.Round_Up_as_double(z1, 2);
                        if (z1 < min1 + 6) min1 = Functions.Round_Down_as_double(z1, 6);
                    }
                }
            }
            if (prof_idx == 1)
            {
                if (max1 > -100000 && min1 < 100000)
                {
                    textBox_el_bottom1.Text = Functions.Get_String_Rounded(min1, 0);
                    textBox_el_top1.Text = Functions.Get_String_Rounded(max1, 0);
                }
            }

            if (prof_idx == 2)
            {
                if (max1 > -100000 && min1 < 100000)
                {
                    textBox_el_bottom2.Text = Functions.Get_String_Rounded(min1, 0);
                    textBox_el_top2.Text = Functions.Get_String_Rounded(max1, 0);
                }
            }

            if (prof_idx == 3)
            {
                if (max1 > -100000 && min1 < 100000)
                {
                    textBox_el_bottom3.Text = Functions.Get_String_Rounded(min1, 0);
                    textBox_el_top3.Text = Functions.Get_String_Rounded(max1, 0);
                }
            }


            return dt1;
        }

        private Polyline Build_2dpoly_from_3d(Polyline3d Poly3D, TextBox textbox1, TextBox textbox2, double multiple_down, double multiple_up)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Polyline Poly2D = new Polyline();
                    int Index1 = 0;
                    if (Poly3D.Length > 0)
                    {

                        double last_param = Poly3D.EndParam;

                        double min_el = 100000;
                        double max_el = -100000;

                        for (int i = 0; i <= last_param; ++i)
                        {
                            try
                            {
                                double z = Poly3D.GetPointAtParameter(i).Z;
                                if (z > max_el) max_el = z;
                                if (z < min_el) min_el = z;
                                Poly2D.AddVertexAt(Index1, new Point2d(Poly3D.GetPointAtParameter(i).X, Poly3D.GetPointAtParameter(i).Y), 0, 0, 0);
                                Index1 = Index1 + 1;

                            }
                            catch (System.Exception ex)
                            {

                            }
                        }

                        if (textbox1 != null) textbox1.Text = Convert.ToString(Functions.Round_Down_as_double(min_el, multiple_down) - multiple_down);
                        if (textbox2 != null) textbox2.Text = Convert.ToString(Functions.Round_Up_as_double(max_el, multiple_up) + multiple_up);

                    }
                    return Poly2D;
                }
            }
        }

        private void Button_Load_stream_cl_Click(object sender, EventArgs e)
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
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a 3D(2D) polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            label_pipe.Visible = false;
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            dt_top_of_bank_ne = null;
                            dt_top_of_bank_sw = null;
                            dt_stream = null;
                            set_enable_true();
                            refcl = null;
                            label_tob_down.Visible = false;
                            label_tob_up.Visible = false;
                            pipe_sta = -1.234;
                            textBox_cl_sta.Text = "";
                            label_int.Visible = false;
                            lista_lod_sta = null;
                            label_LOD.Visible = false;

                            return;
                        }
                        Polyline3d p3 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;

                        if (p3 == null)
                        {
                            Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                            if (p2 != null)
                            {

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez_pt;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the points:";
                                Prompt_rez.SingleOnly = false;
                                rez_pt = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (rez_pt.Status != PromptStatus.OK)
                                {
                                    label_pipe.Visible = false;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    dt_top_of_bank_ne = null;
                                    dt_top_of_bank_sw = null;
                                    dt_stream = null;
                                    set_enable_true();
                                    refcl = null;
                                    label_tob_down.Visible = false;
                                    label_tob_up.Visible = false;
                                    pipe_sta = -1.234;
                                    textBox_cl_sta.Text = "";
                                    label_int.Visible = false;
                                    lista_lod_sta = null;
                                    label_LOD.Visible = false;

                                    return;
                                }

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("sta", typeof(double));
                                dt2.Columns.Add("elev", typeof(double));
                                dt2.Columns.Add("pt", typeof(Point3d));
                                for (int i = 0; i < rez_pt.Value.Count; ++i)
                                {
                                    DBPoint pt1 = Trans1.GetObject(rez_pt.Value[i].ObjectId, OpenMode.ForRead) as DBPoint;
                                    if (pt1 != null)
                                    {
                                        Point3d pt_on_p2 = p2.GetClosestPointTo(new Point3d(pt1.Position.X, pt1.Position.Y, p2.Elevation), Vector3d.ZAxis, false);
                                        double dist2 = Math.Pow(Math.Pow(pt1.Position.X - pt_on_p2.X, 2) + Math.Pow(pt1.Position.Y - pt_on_p2.Y, 2), 0.5);
                                        if (dist2 < 1)
                                        {
                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1]["sta"] = p2.GetDistAtPoint(pt_on_p2);
                                            dt2.Rows[dt2.Rows.Count - 1]["elev"] = pt1.Position.Z;
                                            dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(pt_on_p2.X, pt_on_p2.Y, pt1.Position.Z);

                                        }

                                    }
                                }

                                if (dt2.Rows.Count > 1)
                                {
                                    dt2 = Functions.Sort_data_table(dt2, "sta");
                                    BTrecord.UpgradeOpen();
                                    p3 = new Polyline3d();
                                    p3.Layer = p2.Layer;
                                    BTrecord.AppendEntity(p3);
                                    Trans1.AddNewlyCreatedDBObject(p3, true);

                                    #region first vertex
                                    if (dt2.Rows.Count > 2)
                                    {
                                        double sta1 = Convert.ToDouble(dt2.Rows[0]["sta"]);
                                        if (sta1 > 0)
                                        {
                                            double el1 = Convert.ToDouble(dt2.Rows[0]["elev"]);
                                            double el2 = Convert.ToDouble(dt2.Rows[1]["elev"]);
                                            double sta2 = Convert.ToDouble(dt2.Rows[1]["sta"]);

                                            double z0 = el1 - (((sta1 * (el2 - el1)) / (sta2 - sta1)));

                                            PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(p2.StartPoint.X, p2.StartPoint.Y, z0));
                                            p3.AppendVertex(Vertex_new);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                        }
                                    }


                                    #endregion

                                    for (int i = 1; i < dt2.Rows.Count; ++i)
                                    {
                                        double sta0 = Convert.ToDouble(dt2.Rows[i - 1]["sta"]);
                                        double sta1 = Convert.ToDouble(dt2.Rows[i]["sta"]);

                                        PolylineVertex3d Vertex_new = new PolylineVertex3d((Point3d)dt2.Rows[i - 1]["pt"]);
                                        p3.AppendVertex(Vertex_new);
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                        for (int j = 0; j < p2.NumberOfVertices; ++j)
                                        {
                                            double sta2d = p2.GetDistanceAtParameter(j);
                                            if (sta2d > sta0 && sta2d < sta1)
                                            {
                                                double el0 = Convert.ToDouble(dt2.Rows[i - 1]["elev"]);
                                                double el1 = Convert.ToDouble(dt2.Rows[i]["elev"]);
                                                double elX = el0 + (el1 - el0) * (sta2d - sta0) / (sta1 - sta0);
                                                Vertex_new = new PolylineVertex3d(new Point3d(p2.GetPointAtParameter(j).X, p2.GetPointAtParameter(j).Y, elX));
                                                p3.AppendVertex(Vertex_new);
                                                Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                            }
                                        }


                                        if (i == dt2.Rows.Count - 1)
                                        {
                                            Vertex_new = new PolylineVertex3d((Point3d)dt2.Rows[i]["pt"]);
                                            p3.AppendVertex(Vertex_new);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                        }
                                    }

                                    #region last vertex
                                    if (dt2.Rows.Count > 2)
                                    {
                                        double sta1 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["sta"]);
                                        if (sta1 < p2.Length)
                                        {
                                            double sta_end = p2.Length;
                                            double el1 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["elev"]);
                                            double el2 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 2]["elev"]);
                                            double sta2 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 2]["sta"]);

                                            double z_end = el1 + ((((sta_end - sta1) * (el1 - el2)) / (sta1 - sta2)));

                                            PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(p2.EndPoint.X, p2.EndPoint.Y, z_end));
                                            p3.AppendVertex(Vertex_new);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                        }
                                    }


                                    #endregion

                                    p2.UpgradeOpen();
                                    p2.Erase();
                                }
                            }
                        }


                        if (p3 != null)
                        {

                            refcl = Build_2dpoly_from_3d(p3, textBox_el_bottom1, textBox_el_top1, 6, 2);
                            dt_stream = creaza_data_table(p3, refcl, 1);
                            label_pipe.Visible = true;
                        }
                        else
                        {

                            dt_stream = null;
                            label_pipe.Visible = false;
                            refcl = null;
                        }

                        dt_top_of_bank_ne = null;
                        dt_top_of_bank_sw = null;
                        label_tob_down.Visible = false;
                        label_tob_up.Visible = false;
                        pipe_sta = -1.234;
                        textBox_cl_sta.Text = "";
                        label_int.Visible = false;
                        lista_lod_sta = null;
                        label_LOD.Visible = false;
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
            set_enable_true();
        }

        private void button_load_TOB_ne_Click(object sender, EventArgs e)
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
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the top of bank [north//east]:");
                        Prompt_centerline.SetRejectMessage("\nSelect a 3D polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            label_tob_up.Visible = false;
                            dt_top_of_bank_ne = null;
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }
                        Polyline3d p3 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;
                        if (p3 != null && refcl != null)
                        {
                            dt_top_of_bank_ne = creaza_data_table(p3, refcl, 1);
                            label_tob_up.Visible = true;

                        }
                        else
                        {

                            dt_top_of_bank_ne = null;
                            label_tob_up.Visible = false;

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
            set_enable_true();
        }

        private void button_load_TOB_sw_Click(object sender, EventArgs e)
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
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the top of bank [south//west]:");
                        Prompt_centerline.SetRejectMessage("\nSelect a 3D polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            label_tob_down.Visible = false;
                            dt_top_of_bank_sw = null;
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }

                        Polyline3d p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;
                        if (p2 != null && refcl != null)
                        {
                            dt_top_of_bank_sw = creaza_data_table(p2, refcl, 1);
                            label_tob_down.Visible = true;
                        }
                        else
                        {
                            dt_top_of_bank_sw = null;
                            label_tob_down.Visible = false;
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
            set_enable_true();
        }

        private double get_text_height()
        {
            double nr = 0.08;
            string scale1 = Combobox_scales.Text;
            double s = 1;
            if (scale1.Contains(":") == true)
            {
                int i = scale1.IndexOf(":");
                scale1 = scale1.Substring(i + 1, scale1.Length - i - 1);
                if (Functions.IsNumeric(scale1) == true)
                {
                    s = Convert.ToDouble(scale1);
                }

            }

            return nr * s;
        }

        private void Button_draw_prof_streamcl_Click(object sender, EventArgs e)
        {

            if ((dt_stream != null && dt_stream.Rows.Count > 0))
            {
                if (Functions.IsNumeric(textBox_el_top1.Text) == false)
                {
                    MessageBox.Show("please specify the top elevation1");
                    return;
                }
                if (Functions.IsNumeric(textBox_el_bottom1.Text) == false)
                {
                    MessageBox.Show("please specify the bottom elevation1");
                    return;
                }
            }

            if ((dt_cross != null && dt_cross.Rows.Count > 0))
            {
                if (Functions.IsNumeric(textBox_el_top2.Text) == false)
                {
                    MessageBox.Show("please specify the top elevation2");
                    return;
                }
                if (Functions.IsNumeric(textBox_el_bottom2.Text) == false)
                {
                    MessageBox.Show("please specify the bottom elevation2");
                    return;
                }
            }

            if ((dt_eq != null && dt_eq.Rows.Count > 0))
            {
                if (Functions.IsNumeric(textBox_el_top3.Text) == false)
                {
                    MessageBox.Show("please specify the top elevation3");
                    return;
                }
                if (Functions.IsNumeric(textBox_el_bottom3.Text) == false)
                {
                    MessageBox.Show("please specify the bottom elevation3");
                    return;
                }
            }

            if (Functions.IsNumeric(textBox_prof_Vex.Text) == false)
            {
                MessageBox.Show("please specify the vertical exaggeration");
                return;
            }
            if (Functions.IsNumeric(textBox_prof_Hex.Text) == false)
            {
                MessageBox.Show("please specify the HORIZONTAL exaggeration");
                return;
            }
            if (Functions.IsNumeric(textBox_prof_Vspacing.Text) == false)
            {
                MessageBox.Show("please specify the vertical spacing");
                return;
            }
            if (Functions.IsNumeric(textBox_prof_Hspacing.Text) == false)
            {
                MessageBox.Show("please specify the horizontal spacing");
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


                        double vincr = Convert.ToDouble(textBox_prof_Vspacing.Text);
                        double hincr = Convert.ToDouble(textBox_prof_Hspacing.Text);
                        double textH = get_text_height();
                        string layer_grid_lines = "_pgen_GRID";
                        string layer_text = "_pgen_TEXT";



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify profile  starting point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }

                        Point3d pt_start = Point_res1.Value;

                        draw_profiles(pt_start, hincr, vincr, Convert.ToDouble(textBox_prof_Hex.Text), Convert.ToDouble(textBox_prof_Vex.Text), layer_grid_lines, layer_text, textH, Functions.Get_textstyle_id("Standard"), "'", true, true, "f");

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
            set_enable_true();


        }

        public void draw_profiles(Point3d Point0, double Hincr, double Vincr, double Hexag, double Vexag, string Layer_grid, string Layer_text, double Texth, ObjectId Textstyleid, string Elev_suffix, bool leftElev, bool rightElev, string units)
        {

            string layer_ground = "_pgen_GROUND";
            string layer_TOB_E_N = "_pgen_TOB_E_N";
            string layer_TOB_S_W = "_pgen_TOB_W_S";
            string layer_elipsa = "_pgen_pipe_symbol";
            string layer_lod = "_pgen_LOD";

            Functions.Creaza_layer(Layer_grid, 9, true);
            Functions.Creaza_layer(Layer_text, 2, true);
            Functions.Creaza_layer("no_plot", 30, false);


            Functions.Creaza_layer(layer_ground, 2, true);
            if (dt_top_of_bank_ne != null && dt_top_of_bank_ne.Rows.Count > 1) Functions.Creaza_layer(layer_TOB_E_N, 3, true);
            if (dt_top_of_bank_sw != null && dt_top_of_bank_sw.Rows.Count > 1) Functions.Creaza_layer(layer_TOB_S_W, 3, true);
            if ((dt_eq != null && dt_eq.Rows.Count > 0) || (dt_cross != null && dt_cross.Rows.Count > 0) || textBox_cl_sta.Text != "") Functions.Creaza_layer(layer_elipsa, 1, true);
            if (lista_lod_sta != null && lista_lod_sta.Count == 2) Functions.Creaza_layer(layer_lod, 1, true);


            double Startsta = 0;
            double Endsta = 0;
            double Textwidth = 0;

            double XR = Point0.X;

            string Col_sta = "sta";


            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                #region DT_STREAM
                if ((dt_stream != null && dt_stream.Rows.Count > 0))
                {
                    double Min_sta = 0;
                    double Max_sta = 0;

                    if (dt_stream != null && dt_stream.Rows.Count > 0)
                    {
                        if (dt_stream.Rows[0][Col_sta] != DBNull.Value)
                        {
                            Min_sta = Convert.ToDouble(dt_stream.Rows[0][Col_sta]);
                        }

                        if (dt_stream.Rows[dt_stream.Rows.Count - 1][Col_sta] != DBNull.Value)
                        {
                            Max_sta = Convert.ToDouble(dt_stream.Rows[dt_stream.Rows.Count - 1][Col_sta]);
                        }
                    }



                    Startsta = Functions.Round_Down_as_double(Min_sta, Hincr);
                    Endsta = Functions.Round_Up_as_double(Max_sta, Hincr);

                    double Upelev = Convert.ToDouble(textBox_el_top1.Text);
                    double Downelev = Convert.ToDouble(textBox_el_bottom1.Text);

                    int Nr_linii_elevation = Convert.ToInt32(((Upelev - Downelev) / Vincr) + 1);
                    int Nr_linii_station = Convert.ToInt32(((Endsta - Startsta) / Hincr) + 1);

                    double EndX = Point0.X + (Endsta - Startsta) * Hexag;


                    #region station lines

                    for (int i = 0; i < Nr_linii_station; ++i)
                    {

                        double DisplaySTA = Startsta + i * Hincr;
                        double PozX = i * Hincr * Hexag;


                        Autodesk.AutoCAD.DatabaseServices.Line LinieV = new Autodesk.AutoCAD.DatabaseServices.Line(
                                                                                          new Point3d(Point0.X + PozX, Point0.Y, 0),
                                                                                          new Point3d(Point0.X + PozX, Point0.Y + (Upelev - Downelev) * Vexag, 0));
                        LinieV.Layer = Layer_grid;
                        LinieV.Linetype = "ByLayer";
                        LinieV.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieV);
                        Trans1.AddNewlyCreatedDBObject(LinieV, true);

                        MText Mt_sta = new MText();
                        Mt_sta.Contents = Functions.Get_chainage_from_double(DisplaySTA, units, 0);
                        Mt_sta.Layer = Layer_text;
                        Mt_sta.Attachment = AttachmentPoint.TopCenter;
                        Mt_sta.TextHeight = Texth;
                        Mt_sta.TextStyleId = Textstyleid;
                        Mt_sta.Location = new Point3d(Point0.X + PozX, Point0.Y - 2 * Texth, 0);
                        Mt_sta.ColorIndex = 256;
                        BTrecord.AppendEntity(Mt_sta);
                        Trans1.AddNewlyCreatedDBObject(Mt_sta, true);
                    }

                    #endregion

                    #region cl line
                    Line Linie_cl = null;
                    if (pipe_sta != -1.234)
                    {
                        Linie_cl = new Line(new Point3d(Point0.X + (pipe_sta - Startsta) * Hexag, Point0.Y, 0), new Point3d(Point0.X + +(pipe_sta - Startsta) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + 2 * Vincr * Vexag, 0));
                        Linie_cl.Layer = "no_plot";
                        Linie_cl.Linetype = "ByLayer";
                        Linie_cl.ColorIndex = 256;
                        BTrecord.AppendEntity(Linie_cl);
                        Trans1.AddNewlyCreatedDBObject(Linie_cl, true);
                    }

                    #endregion





                    #region elevation lines

                    bool draw_left = false;
                    bool draw_right = false;

                    for (int i = 0; i < Nr_linii_elevation; ++i)
                    {

                        Autodesk.AutoCAD.DatabaseServices.Line LinieH =
                            new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(Point0.X, Point0.Y + i * Vincr * Vexag, 0),
                                                                       new Point3d(EndX, Point0.Y + i * Vincr * Vexag, 0));

                        LinieH.Layer = Layer_grid;
                        LinieH.Linetype = "ByLayer";
                        LinieH.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieH);
                        Trans1.AddNewlyCreatedDBObject(LinieH, true);

                        if (leftElev == true)
                        {
                            if (checkBox_label_half.Checked == false)
                            {
                                draw_left = true;
                            }
                            else
                            {
                                if (draw_left == true)
                                {
                                    draw_left = false;
                                }
                                else
                                {
                                    draw_left = true;
                                }
                            }
                        }

                        if (draw_left == true)
                        {
                            MText Mt_el_left = new MText();
                            Mt_el_left.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_left.Layer = Layer_text;
                            Mt_el_left.Attachment = AttachmentPoint.MiddleRight;
                            Mt_el_left.TextHeight = Texth;
                            Mt_el_left.TextStyleId = Textstyleid;
                            Mt_el_left.Location = new Point3d(Point0.X - 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_left.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_left);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_left, true);

                            Extents3d Extend1 = Mt_el_left.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Extend1.MaxPoint.X - Extend1.MinPoint.X;
                            }
                        }


                        if (rightElev == true)
                        {
                            if (checkBox_label_half.Checked == false)
                            {
                                draw_right = true;
                            }
                            else
                            {
                                if (draw_right == true)
                                {
                                    draw_right = false;
                                }
                                else
                                {
                                    draw_right = true;
                                }
                            }
                        }

                        if (draw_right == true)
                        {

                            MText Mt_el_right = new MText();
                            Mt_el_right.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_right.Layer = Layer_text;
                            Mt_el_right.Attachment = AttachmentPoint.MiddleLeft;
                            Mt_el_right.TextHeight = Texth;
                            Mt_el_right.TextStyleId = Textstyleid;
                            Mt_el_right.Location = new Point3d(EndX + 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_right.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_right);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_right, true);

                            XR = EndX + 2 * Texth;

                            Extents3d Extend1 = Mt_el_right.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Math.Abs(Extend1.MaxPoint.X - Extend1.MinPoint.X);
                            }

                        }
                    }

                    #endregion


                    #region poly graphs
                    Polyline Poly_graph1 = new Polyline();
                    int idx_p = 0;


                    if (dt_stream != null && dt_stream.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_stream.Rows.Count; ++i)
                        {
                            if (dt_stream.Rows[i]["z"] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_stream.Rows[i]["z"]);
                                if (dt_stream.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_stream.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph1.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph1.Plinegen = true;
                        Poly_graph1.Layer = layer_ground;
                        Poly_graph1.ColorIndex = 256;
                        BTrecord.AppendEntity(Poly_graph1);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph1, true);
                    }



                    Polyline Poly_graph2 = new Polyline();
                    idx_p = 0;

                    if (dt_top_of_bank_ne != null && dt_top_of_bank_ne.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_top_of_bank_ne.Rows.Count; ++i)
                        {
                            if (dt_top_of_bank_ne.Rows[i]["z"] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_top_of_bank_ne.Rows[i]["z"]);
                                if (dt_top_of_bank_ne.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_top_of_bank_ne.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph2.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph2.Plinegen = true;
                        Poly_graph2.Layer = layer_TOB_E_N;
                        Poly_graph2.ColorIndex = 256;
                        BTrecord.AppendEntity(Poly_graph2);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph2, true);
                    }

                    Polyline Poly_graph3 = new Polyline();
                    idx_p = 0;

                    if (dt_top_of_bank_sw != null && dt_top_of_bank_sw.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_top_of_bank_sw.Rows.Count; ++i)
                        {
                            if (dt_top_of_bank_sw.Rows[i]["z"] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_top_of_bank_sw.Rows[i]["z"]);
                                if (dt_top_of_bank_sw.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_top_of_bank_sw.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph3.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph3.Plinegen = true;
                        Poly_graph3.Layer = layer_TOB_S_W;
                        Poly_graph3.ColorIndex = 256;
                        BTrecord.AppendEntity(Poly_graph3);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph3, true);
                    }



                    #endregion


                    #region lod
                    if (lista_lod_sta != null && lista_lod_sta.Count > 0)
                    {
                        Point3dCollection colint = new Point3dCollection();

                        for (int i = 0; i < lista_lod_sta.Count; ++i)
                        {
                            Line Linie_loD = new Line(new Point3d(Point0.X + (lista_lod_sta[i] - Startsta) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (lista_lod_sta[i] - Startsta) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                            Linie_loD.Layer = "no_plot";
                            Linie_loD.Linetype = "ByLayer";
                            Linie_loD.ColorIndex = 256;
                            BTrecord.AppendEntity(Linie_loD);
                            Trans1.AddNewlyCreatedDBObject(Linie_loD, true);
                            Point3dCollection col3 = Functions.Intersect_on_both_operands(Linie_loD, Poly_graph1);
                            if (col3.Count > 0)
                            {
                                colint.Add(col3[0]);
                            }

                        }

                        if (lista_lod_sta.Count == 2 && colint.Count == 2)
                        {
                            double sta1 = lista_lod_sta[0];
                            double sta2 = lista_lod_sta[1];

                            Point3d pt1 = colint[0];
                            Point3d pt2 = colint[1];

                            if (sta1 > sta2)
                            {
                                double t = sta1;
                                sta1 = sta2;
                                sta2 = t;

                                Point3d ptt = pt1;
                                pt1 = pt2;
                                pt2 = ptt;

                            }



                            Point3d ptm = new Point3d((pt1.X + pt2.X) / 2, 2 * Vincr * Vexag + Point0.Y + (Upelev - Downelev) * Vexag, 0);

                            RotatedDimension dim1 = new RotatedDimension();
                            dim1.XLine1Point = pt1;
                            dim1.XLine2Point = pt2;
                            dim1.Rotation = 0;
                            dim1.DimLinePoint = ptm;
                            dim1.Dimasz = Texth;
                            dim1.Dimtxt = Texth;
                            dim1.Layer = layer_lod;
                            dim1.HorizontalRotation = 0;
                            dim1.Dimtfill = 1;
                            dim1.Dimdec = 0;
                            dim1.Dimsd1 = false;
                            dim1.Dimsd2 = false;
                            dim1.Dimse1 = false;
                            dim1.Dimse2 = false;
                            dim1.Dimlfac = 1 / Vexag;
                            dim1.Dimexe = Texth;
                            dim1.DimensionText = "LIMITS OF DISTURBANCE";

                            BTrecord.AppendEntity(dim1);
                            Trans1.AddNewlyCreatedDBObject(dim1, true);


                        }

                    }

                    #endregion

                    #region mleader Streambed
                    if (dt_stream != null && dt_stream.Rows.Count > 1)
                    {
                        Line Linie_1 = new Line(new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_1, Poly_graph1);
                        if (colint.Count > 0)
                        {
                            Point3d pt1 = colint[0];
                            Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                            MLeader streambed = creaza_mleader(pt1, "STREAMBED \u2104", Texth, 6 * Texth, -8 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                        }

                    }
                    #endregion

                    #region mleader TOP OF EAST BANK
                    if (dt_top_of_bank_ne != null && dt_top_of_bank_ne.Rows.Count > 1)
                    {
                        Line Linie_1 = new Line(new Point3d(Point0.X + (((Endsta + Startsta) / 2) - 2 * Hincr) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (((Endsta + Startsta) / 2) - 2 * Hincr) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_1, Poly_graph2);
                        if (colint.Count > 0)
                        {
                            Point3d pt1 = colint[0];
                            Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                            MLeader streambed = creaza_mleader(pt1, "TOP OF EAST BANK", Texth, 6 * Texth, 8 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                        }

                    }
                    #endregion


                    #region mleader TOP OF WEST BANK
                    if (dt_top_of_bank_sw != null && dt_top_of_bank_sw.Rows.Count > 1)
                    {
                        Line Linie_1 = new Line(new Point3d(Point0.X + (((Endsta + Startsta) / 2) - 2 * Hincr) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (((Endsta + Startsta) / 2) - 2 * Hincr) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_1, Poly_graph3);
                        if (colint.Count > 0)
                        {
                            Point3d pt1 = colint[0];
                            Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                            MLeader streambed = creaza_mleader(pt1, "TOP OF WEST BANK", Texth, 6 * Texth, 8 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                        }

                    }
                    #endregion


                    #region pipe symbol
                    if (dt_stream != null && dt_stream.Rows.Count > 1 && Linie_cl != null)
                    {

                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_cl, Poly_graph1);
                        if (colint.Count > 0 && Functions.IsNumeric(textBox_pipe_diam.Text) == true)
                        {

                            double diam_inch = Convert.ToDouble(textBox_pipe_diam.Text);
                            Point3d pt1 = colint[0];

                            BlockTable1.UpgradeOpen();

                            int idx = 1;
                            bool exista = true;
                            do
                            {
                                if (BlockTable1.Has("p_sym" + idx.ToString()) == false)
                                {
                                    using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                    {
                                        bltrec1.Name = "p_sym" + idx.ToString();
                                        Circle cerc1 = new Circle(new Point3d(0, -0.5 * diam_inch / 12, 0), Vector3d.ZAxis, 0.5 * diam_inch / 12);

                                        cerc1.Layer = "0";
                                        cerc1.ColorIndex = 0;
                                        bltrec1.AppendEntity(cerc1);
                                        BlockTable1.Add(bltrec1);
                                        Trans1.AddNewlyCreatedDBObject(bltrec1, true);
                                    }
                                    exista = false;
                                }
                                else
                                {
                                    ++idx;
                                }
                            } while (exista == true);

                            Point3d ptins = new Point3d(pt1.X, pt1.Y - 5 * Vexag, 0);


                            InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, "p_sym" + idx.ToString(), ptins, Hexag, Vexag, 0, layer_elipsa);

                            Point3d ptm = new Point3d((pt1.X + ptins.X) / 2, (pt1.Y + ptins.Y) / 2, 0);

                            RotatedDimension dim1 = new RotatedDimension();
                            dim1.XLine1Point = ptins;
                            dim1.XLine2Point = pt1;
                            dim1.Rotation = Math.PI / 2;
                            dim1.DimLinePoint = ptm;
                            dim1.Dimasz = Texth;
                            dim1.Dimtxt = Texth;
                            dim1.Layer = layer_elipsa;
                            dim1.HorizontalRotation = -Math.PI / 2;
                            dim1.Dimtfill = 1;
                            dim1.Dimdec = 0;
                            dim1.Dimsd1 = false;
                            dim1.Dimsd2 = false;
                            dim1.Dimse1 = true;
                            dim1.Dimse2 = true;
                            dim1.Dimlfac = 1 / Vexag;
                            dim1.DimensionText = "<>' (MIN)";

                            BTrecord.AppendEntity(dim1);
                            Trans1.AddNewlyCreatedDBObject(dim1, true);

                            Ellipse elipsa1 = new Ellipse(new Point3d(ptins.X, ptins.Y - Vexag * 0.5 * diam_inch / 12, 0), Vector3d.ZAxis, new Vector3d(0, Vexag * 0.5 * diam_inch / 12, 0), 0.5, 0, 2 * Math.PI);
                            Line lin1 = new Line(elipsa1.Center, new Point3d(elipsa1.Center.X + diam_inch, elipsa1.Center.Y, 0));
                            lin1.TransformBy(Matrix3d.Rotation(70 * Math.PI / 180, Vector3d.ZAxis, elipsa1.Center));
                            Point3dCollection col2 = Functions.Intersect_on_both_operands(elipsa1, lin1);
                            if (col2.Count > 0)
                            {
                                Point3d pt2 = col2[0];
                                Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                                MLeader pipemleader = creaza_mleader(pt2, textBox1.Text + " " + textBox_pipe_diam.Text + textBox2.Text + "\r\n" + textBox3.Text, Texth, 4 * Texth, 6 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                            }



                        }

                    }
                    #endregion

                    Point0 = new Point3d(EndX + 100, Point0.Y, 0);
                    XR = Point0.X;

                }
                #endregion


                #region DT_cross
                if ((dt_cross != null && dt_cross.Rows.Count > 0))
                {
                    Functions.Creaza_layer("PGEN_MTEXT", 1, true);

                    double Min_sta = 0;
                    double Max_sta = 0;

                    if (dt_cross != null && dt_cross.Rows.Count > 0)
                    {
                        if (dt_cross.Rows[0][Col_sta] != DBNull.Value)
                        {
                            Min_sta = Convert.ToDouble(dt_cross.Rows[0][Col_sta]);
                        }

                        if (dt_cross.Rows[dt_cross.Rows.Count - 1][Col_sta] != DBNull.Value)
                        {
                            Max_sta = Convert.ToDouble(dt_cross.Rows[dt_cross.Rows.Count - 1][Col_sta]);
                        }
                    }



                    Startsta = Functions.Round_Down_as_double(Min_sta, Hincr);
                    Endsta = Functions.Round_Up_as_double(Max_sta, Hincr);

                    double Upelev = Convert.ToDouble(textBox_el_top2.Text);
                    double Downelev = Convert.ToDouble(textBox_el_bottom2.Text);

                    int Nr_linii_elevation = Convert.ToInt32(((Upelev - Downelev) / Vincr) + 1);
                    int Nr_linii_station = Convert.ToInt32(((Endsta - Startsta) / Hincr) + 1);

                    double EndX = Point0.X + (Endsta - Startsta) * Hexag;


                    #region station lines

                    for (int i = 0; i < Nr_linii_station; ++i)
                    {

                        double DisplaySTA = Startsta + i * Hincr;
                        double PozX = i * Hincr * Hexag;


                        Autodesk.AutoCAD.DatabaseServices.Line LinieV = new Autodesk.AutoCAD.DatabaseServices.Line(
                                                                                          new Point3d(Point0.X + PozX, Point0.Y, 0),
                                                                                          new Point3d(Point0.X + PozX, Point0.Y + (Upelev - Downelev) * Vexag, 0));
                        LinieV.Layer = Layer_grid;
                        LinieV.Linetype = "ByLayer";
                        LinieV.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieV);
                        Trans1.AddNewlyCreatedDBObject(LinieV, true);

                        MText Mt_sta = new MText();
                        Mt_sta.Contents = Functions.Get_chainage_from_double(DisplaySTA, units, 0);
                        Mt_sta.Layer = Layer_text;
                        Mt_sta.Attachment = AttachmentPoint.TopCenter;
                        Mt_sta.TextHeight = Texth;
                        Mt_sta.TextStyleId = Textstyleid;
                        Mt_sta.Location = new Point3d(Point0.X + PozX, Point0.Y - 2 * Texth, 0);
                        Mt_sta.ColorIndex = 256;
                        BTrecord.AppendEntity(Mt_sta);
                        Trans1.AddNewlyCreatedDBObject(Mt_sta, true);
                    }

                    #endregion




                    #region elevation lines
                    for (int i = 0; i < Nr_linii_elevation; ++i)
                    {

                        Autodesk.AutoCAD.DatabaseServices.Line LinieH =
                            new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(Point0.X, Point0.Y + i * Vincr * Vexag, 0),
                                                                       new Point3d(EndX, Point0.Y + i * Vincr * Vexag, 0));

                        LinieH.Layer = Layer_grid;
                        LinieH.Linetype = "ByLayer";
                        LinieH.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieH);
                        Trans1.AddNewlyCreatedDBObject(LinieH, true);

                        if (leftElev == true)
                        {
                            MText Mt_el_left = new MText();
                            Mt_el_left.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_left.Layer = Layer_text;
                            Mt_el_left.Attachment = AttachmentPoint.MiddleRight;
                            Mt_el_left.TextHeight = Texth;
                            Mt_el_left.TextStyleId = Textstyleid;
                            Mt_el_left.Location = new Point3d(Point0.X - 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_left.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_left);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_left, true);

                            Extents3d Extend1 = Mt_el_left.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Extend1.MaxPoint.X - Extend1.MinPoint.X;
                            }

                        }

                        if (rightElev == true)
                        {
                            MText Mt_el_right = new MText();
                            Mt_el_right.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_right.Layer = Layer_text;
                            Mt_el_right.Attachment = AttachmentPoint.MiddleLeft;
                            Mt_el_right.TextHeight = Texth;
                            Mt_el_right.TextStyleId = Textstyleid;
                            Mt_el_right.Location = new Point3d(EndX + 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_right.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_right);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_right, true);

                            XR = EndX + 2 * Texth;

                            Extents3d Extend1 = Mt_el_right.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Math.Abs(Extend1.MaxPoint.X - Extend1.MinPoint.X);
                            }

                        }
                    }

                    #endregion


                    #region poly graph
                    Polyline Poly_graph1 = new Polyline();
                    int idx_p = 0;

                    Point3d min_el_pt = new Point3d();

                    double z_min = +100000;

                    if (dt_cross != null && dt_cross.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_cross.Rows.Count; ++i)
                        {
                            if (dt_cross.Rows[i]["z"] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_cross.Rows[i]["z"]);
                                if (dt_cross.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_cross.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph1.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;

                                    if (z1 < z_min)
                                    {
                                        z_min = z1;
                                        min_el_pt = new Point3d(ptp.X, ptp.Y, 0);
                                    }

                                }
                            }
                        }
                        Poly_graph1.Plinegen = true;
                        Poly_graph1.Layer = layer_ground;
                        Poly_graph1.ColorIndex = 256;
                        BTrecord.AppendEntity(Poly_graph1);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph1, true);
                    }







                    #endregion



                    #region mleader Existing Grade
                    if (dt_cross != null && dt_cross.Rows.Count > 1)
                    {
                        Line Linie_1 = new Line(new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_1, Poly_graph1);
                        if (colint.Count > 0)
                        {
                            Point3d pt1 = colint[0];

                            Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                            MLeader streambed = creaza_mleader(pt1, "EXISTING GRADE", Texth, 6 * Texth, 8 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                        }

                    }
                    #endregion


                    #region block pipe

                    InsertBlock_with_2scales(ThisDrawing.Database, BTrecord, "BELOW_GRADE_PIPE", min_el_pt, 1, 1, 0, layer_elipsa);

                    Point3d pt_low = new Point3d(min_el_pt.X, min_el_pt.Y - 5 * Vexag, 0);
                    Point3d pt_high = new Point3d(min_el_pt.X, min_el_pt.Y, 0);
                    Point3d ptm = new Point3d((pt_low.X + pt_high.X) / 2, (pt_low.Y + pt_high.Y) / 2, 0);

                    RotatedDimension dim1 = new RotatedDimension();
                    dim1.XLine1Point = pt_low;
                    dim1.XLine2Point = pt_high;
                    dim1.Rotation = Math.PI / 2;
                    dim1.DimLinePoint = ptm;
                    dim1.Dimasz = Texth;
                    dim1.Dimtxt = Texth;
                    dim1.Layer = layer_elipsa;
                    dim1.HorizontalRotation = 0;
                    dim1.Dimtfill = 1;
                    dim1.Dimdec = 0;
                    dim1.Dimsd1 = false;
                    dim1.Dimsd2 = false;
                    dim1.Dimse1 = true;
                    dim1.Dimse2 = true;
                    dim1.Dimpost = "";
                    dim1.Dimlfac = 1 / Vexag;
                    dim1.DimensionText = "<>' (MIN)";
                    dim1.ColorIndex = 256;

                    BTrecord.AppendEntity(dim1);
                    Trans1.AddNewlyCreatedDBObject(dim1, true);

                    Point3d pt_ins_ml = new Point3d(pt_low.X + Hincr * Hexag, pt_low.Y, 0);
                    MLeader pipemleader = creaza_mleader(pt_ins_ml, textBox1.Text + " " + textBox_pipe_diam.Text + textBox2.Text + "\r\n" + textBox3.Text, Texth, 4 * Texth, 6 * Texth, Texth, Texth, Texth, layer_elipsa);


                    #endregion


                    #region Mtext 



                    MText mtext1 = new MText();
                    mtext1.Attachment = AttachmentPoint.MiddleLeft;
                    mtext1.Contents = "{\\LEL. " + Functions.Get_String_Rounded(z_min, 2) + "' (INV.)}";
                    mtext1.TextHeight = Texth;
                    mtext1.BackgroundFill = true;
                    mtext1.UseBackgroundColor = true;
                    mtext1.BackgroundScaleFactor = 1.2;
                    mtext1.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y, 0);
                    mtext1.Rotation = 0;
                    mtext1.ColorIndex = 256;
                    mtext1.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext1);
                    Trans1.AddNewlyCreatedDBObject(mtext1, true);


                    MText mtext2 = new MText();
                    mtext2.Attachment = AttachmentPoint.MiddleLeft;
                    mtext2.Contents = "{\\LEL. TBD (2YR)}";
                    mtext2.TextHeight = Texth;
                    mtext2.BackgroundFill = true;
                    mtext2.UseBackgroundColor = true;
                    mtext2.BackgroundScaleFactor = 1.2;
                    mtext2.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y + 1.5 * Texth, 0);
                    mtext2.Rotation = 0;
                    mtext2.ColorIndex = 10;
                    mtext2.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext2);
                    Trans1.AddNewlyCreatedDBObject(mtext2, true);

                    MText mtext3 = new MText();
                    mtext3.Attachment = AttachmentPoint.MiddleLeft;
                    mtext3.Contents = "{\\L\\u+25BC 2YR DEPTH}";
                    mtext3.TextHeight = Texth;
                    mtext3.BackgroundFill = true;
                    mtext3.UseBackgroundColor = true;
                    mtext3.BackgroundScaleFactor = 1.2;
                    mtext3.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y - 1.5 * Texth, 0);
                    mtext3.Rotation = 0;
                    mtext3.ColorIndex = 10;
                    mtext3.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext3);
                    Trans1.AddNewlyCreatedDBObject(mtext3, true);


                    #endregion

                    Point0 = new Point3d(EndX + 100, Point0.Y, 0);
                    XR = Point0.X;

                }
                #endregion


                #region DT_eq
                if ((dt_eq != null && dt_eq.Rows.Count > 0))
                {

                    Functions.Creaza_layer("PGEN_MTEXT", 1, true);


                    double Min_sta = 0;
                    double Max_sta = 0;

                    if (dt_eq != null && dt_eq.Rows.Count > 0)
                    {
                        if (dt_eq.Rows[0][Col_sta] != DBNull.Value)
                        {
                            Min_sta = Convert.ToDouble(dt_eq.Rows[0][Col_sta]);
                        }

                        if (dt_eq.Rows[dt_eq.Rows.Count - 1][Col_sta] != DBNull.Value)
                        {
                            Max_sta = Convert.ToDouble(dt_eq.Rows[dt_eq.Rows.Count - 1][Col_sta]);
                        }
                    }



                    Startsta = Functions.Round_Down_as_double(Min_sta, Hincr);
                    Endsta = Functions.Round_Up_as_double(Max_sta, Hincr);

                    double Upelev = Convert.ToDouble(textBox_el_top3.Text);
                    double Downelev = Convert.ToDouble(textBox_el_bottom3.Text);

                    int Nr_linii_elevation = Convert.ToInt32(((Upelev - Downelev) / Vincr) + 1);
                    int Nr_linii_station = Convert.ToInt32(((Endsta - Startsta) / Hincr) + 1);

                    double EndX = Point0.X + (Endsta - Startsta) * Hexag;


                    #region station lines

                    for (int i = 0; i < Nr_linii_station; ++i)
                    {

                        double DisplaySTA = Startsta + i * Hincr;
                        double PozX = i * Hincr * Hexag;


                        Autodesk.AutoCAD.DatabaseServices.Line LinieV = new Autodesk.AutoCAD.DatabaseServices.Line(
                                                                                          new Point3d(Point0.X + PozX, Point0.Y, 0),
                                                                                          new Point3d(Point0.X + PozX, Point0.Y + (Upelev - Downelev) * Vexag, 0));
                        LinieV.Layer = Layer_grid;
                        LinieV.Linetype = "ByLayer";
                        LinieV.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieV);
                        Trans1.AddNewlyCreatedDBObject(LinieV, true);

                        MText Mt_sta = new MText();
                        Mt_sta.Contents = Functions.Get_chainage_from_double(DisplaySTA, units, 0);
                        Mt_sta.Layer = Layer_text;
                        Mt_sta.Attachment = AttachmentPoint.TopCenter;
                        Mt_sta.TextHeight = Texth;
                        Mt_sta.TextStyleId = Textstyleid;
                        Mt_sta.Location = new Point3d(Point0.X + PozX, Point0.Y - 2 * Texth, 0);
                        Mt_sta.ColorIndex = 256;
                        BTrecord.AppendEntity(Mt_sta);
                        Trans1.AddNewlyCreatedDBObject(Mt_sta, true);
                    }

                    #endregion




                    #region elevation lines
                    for (int i = 0; i < Nr_linii_elevation; ++i)
                    {

                        Autodesk.AutoCAD.DatabaseServices.Line LinieH =
                            new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(Point0.X, Point0.Y + i * Vincr * Vexag, 0),
                                                                       new Point3d(EndX, Point0.Y + i * Vincr * Vexag, 0));

                        LinieH.Layer = Layer_grid;
                        LinieH.Linetype = "ByLayer";
                        LinieH.ColorIndex = 256;
                        BTrecord.AppendEntity(LinieH);
                        Trans1.AddNewlyCreatedDBObject(LinieH, true);

                        if (leftElev == true)
                        {
                            MText Mt_el_left = new MText();
                            Mt_el_left.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_left.Layer = Layer_text;
                            Mt_el_left.Attachment = AttachmentPoint.MiddleRight;
                            Mt_el_left.TextHeight = Texth;
                            Mt_el_left.TextStyleId = Textstyleid;
                            Mt_el_left.Location = new Point3d(Point0.X - 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_left.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_left);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_left, true);

                            Extents3d Extend1 = Mt_el_left.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Extend1.MaxPoint.X - Extend1.MinPoint.X;
                            }

                        }

                        if (rightElev == true)
                        {
                            MText Mt_el_right = new MText();
                            Mt_el_right.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_right.Layer = Layer_text;
                            Mt_el_right.Attachment = AttachmentPoint.MiddleLeft;
                            Mt_el_right.TextHeight = Texth;
                            Mt_el_right.TextStyleId = Textstyleid;
                            Mt_el_right.Location = new Point3d(EndX + 2 * Texth, Point0.Y + i * Vincr * Vexag, 0);
                            Mt_el_right.ColorIndex = 256;
                            BTrecord.AppendEntity(Mt_el_right);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_right, true);

                            XR = EndX + 2 * Texth;

                            Extents3d Extend1 = Mt_el_right.GeometricExtents;

                            if (Extend1.MaxPoint.X - Extend1.MinPoint.X > Textwidth)
                            {
                                Textwidth = Math.Abs(Extend1.MaxPoint.X - Extend1.MinPoint.X);
                            }

                        }
                    }

                    #endregion


                    #region poly graph
                    Polyline Poly_graph1 = new Polyline();
                    int idx_p = 0;

                    Point3d min_el_pt = new Point3d();

                    double z_min = +100000;

                    if (dt_eq != null && dt_eq.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_eq.Rows.Count; ++i)
                        {
                            if (dt_eq.Rows[i]["z"] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_eq.Rows[i]["z"]);
                                if (dt_eq.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_eq.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph1.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;

                                    if (z1 < z_min)
                                    {
                                        z_min = z1;
                                        min_el_pt = new Point3d(ptp.X, ptp.Y, 0);
                                    }

                                }
                            }
                        }
                        Poly_graph1.Plinegen = true;
                        Poly_graph1.Layer = layer_ground;
                        Poly_graph1.ColorIndex = 256;
                        BTrecord.AppendEntity(Poly_graph1);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph1, true);
                    }







                    #endregion



                    #region mleader existing grade
                    if (dt_eq != null && dt_eq.Rows.Count > 1)
                    {
                        Line Linie_1 = new Line(new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y, 0), new Point3d(Point0.X + (((Endsta + Startsta) / 2) + 2 * Hincr) * Hexag, Point0.Y + (Upelev - Downelev) * Vexag + Vincr * Vexag, 0));
                        Point3dCollection colint = Functions.Intersect_on_both_operands(Linie_1, Poly_graph1);
                        if (colint.Count > 0)
                        {
                            Point3d pt1 = colint[0];
                            Functions.Creaza_layer("PGEN_MLEADERS", 2, true);
                            MLeader streambed = creaza_mleader(pt1, "EXISTING GRADE", Texth, 6 * Texth, 8 * Texth, Texth, Texth, Texth, "PGEN_MLEADERS");

                        }

                    }
                    #endregion

                    #region Mtext 



                    MText mtext1 = new MText();
                    mtext1.Attachment = AttachmentPoint.MiddleLeft;
                    mtext1.Contents = "{\\LEL. " + Functions.Get_String_Rounded(z_min, 2) + "' (INV.)}";
                    mtext1.TextHeight = Texth;
                    mtext1.BackgroundFill = true;
                    mtext1.UseBackgroundColor = true;
                    mtext1.BackgroundScaleFactor = 1.2;
                    mtext1.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y, 0);
                    mtext1.Rotation = 0;
                    mtext1.ColorIndex = 256;
                    mtext1.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext1);
                    Trans1.AddNewlyCreatedDBObject(mtext1, true);


                    MText mtext2 = new MText();
                    mtext2.Attachment = AttachmentPoint.MiddleLeft;
                    mtext2.Contents = "{\\LEL. TBD (2YR)}";
                    mtext2.TextHeight = Texth;
                    mtext2.BackgroundFill = true;
                    mtext2.UseBackgroundColor = true;
                    mtext2.BackgroundScaleFactor = 1.2;
                    mtext2.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y + 1.5 * Texth, 0);
                    mtext2.Rotation = 0;
                    mtext2.ColorIndex = 10;
                    mtext2.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext2);
                    Trans1.AddNewlyCreatedDBObject(mtext2, true);

                    MText mtext3 = new MText();
                    mtext3.Attachment = AttachmentPoint.MiddleLeft;
                    mtext3.Contents = "{\\L\\u+25BC 2YR DEPTH}";
                    mtext3.TextHeight = Texth;
                    mtext3.BackgroundFill = true;
                    mtext3.UseBackgroundColor = true;
                    mtext3.BackgroundScaleFactor = 1.2;
                    mtext3.Location = new Point3d(min_el_pt.X + 4 * Texth, min_el_pt.Y - 1.5 * Texth, 0);
                    mtext3.Rotation = 0;
                    mtext3.ColorIndex = 10;
                    mtext3.Layer = "PGEN_MTEXT";
                    BTrecord.AppendEntity(mtext3);
                    Trans1.AddNewlyCreatedDBObject(mtext3, true);


                    #endregion


                }
                #endregion


                Trans1.Commit();
            }


        }

        private void dimensions()
        {
            //With Dimension1
            //.Dimasz = 18 'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
            //            'Multiples of the arrowhead size determine whether dimension lines and text should fit between the extension lines. DIMASZ is also used to scale arrowhead blocks if set by DIMBLK. DIMASZ has no effect when DIMTSZ is other than zero

            // .Dimdec = 0
            //            'Sets the number of decimal places displayed for the primary units of a dimension
            //            'The precision is based on the units or angle format you have selected. 


            //  .Dimtxt = 8 'Specifies the height of dimension text, unless the current text style has a fixed height

            //            .TextStyleId = Text_style_romans.ObjectId

            //            .Dimtxtdirection = False
            //            'Specifies the reading direction of the dimension text. 
            //            '0 - Displays dimension text in a Left-to-Right reading style 
            //            '1 - Displays dimension text in a Right-to-Left reading style  



            //            .Dimtofl = False
            //            'Initial value: Off (imperial) or On (metric)  
            //            'Controls whether a dimension line is drawn between the extension lines even when the text is placed outside. 
            //            'For radius and diameter dimensions (when DIMTIX is off), draws a dimension line inside the circle or arc and places the text, arrowheads, and leader outside. 
            //            ' Off -  Does not draw dimension lines between the measured points when arrowheads are placed outside the measured points 
            //            ' On -  Draws dimension lines between the measured points even when arrowheads are placed outside the measured points 

            //            .Dimtoh = False
            //            'Controls the position of dimension text outside the extension lines. 
            //            ' Off -  Aligns text with the dimension line
            //            ' On -  Draws text horizontally

            //            .Dimtih = False
            //            'Initial value: On (imperial) or Off (metric)  
            //            'Controls the position of dimension text inside the extension lines for all dimension types except Ordinate. 
            //            'Off - Aligns text with the dimension line
            //            'On -  Draws text horizontally

            //            .Dimtad = 0
            //            'Controls the vertical position of text in relation to the dimension line. 
            //            '0 - Centers the dimension text between the extension lines. 
            //            '1 - Places the dimension text above the dimension line except when the dimension line is not horizontal and text inside the extension lines is forced horizontal ( DIMTIH = 1). 
            //            '    The distance from the dimension line to the baseline of the lowest line of text is the current DIMGAP value. 
            //            '2 - Places the dimension text on the side of the dimension line farthest away from the defining points. 
            //            '3 - Places the dimension text to conform to Japanese Industrial Standards (JIS). 
            //            '4 - Places the dimension text below the dimension line. 


            //            .Dimtvp = 0
            //            'Controls the vertical position of dimension text above or below the dimension line. 
            //            'The DIMTVP value is used when DIMTAD is off. The magnitude of the vertical offset of text is the product of the text height and DIMTVP. 
            //            'Setting DIMTVP to 1.0 is equivalent to setting DIMTAD to on. The dimension line splits to accommodate the text only if the absolute value of DIMTVP is less than 0.7. 


            //            .Dimsd1 = False
            //            'Controls suppression of the first dimension line and arrowhead. 
            //            'When turned on, suppresses the display of the dimension line and arrowhead between the first extension line and the text. 
            //            .Dimsd2 = False
            //            'Controls suppression of the second dimension line and arrowhead. 
            //            'When turned on, suppresses the display of the dimension line and arrowhead between the second extension line and the text. 
            //            .Dimse1 = True 'Suppresses display of the first extension line. 
            //            .Dimse2 = True 'Suppresses display of the second extension line

            //            .Dimrnd = 5
            //            'Rounds all dimensioning distances to the specified value. 
            //            'For instance, if DIMRND is set to 0.25, all distances round to the nearest 0.25 unit. 
            //            'If you set DIMRND to 1.0, all distances round to the nearest integer. 
            //            'Note that the number of digits edited after the decimal point depends on the precision set by DIMDEC. DIMRND does not apply to angular dimensions. 

            //            .Dimpost = "<>'"
            //            'Specifies a text prefix or suffix (or both) to the dimension measurement. 
            //            'For example, to establish a suffix for millimeters, set DIMPOST to mm; a distance of 19.2 units would be displayed as 19.2 mm. 
            //            'If tolerances are turned on, the suffix is applied to the tolerances as well as to the main dimension. 
            //            'Use <> to indicate placement of the text in relation to the dimension value. 
            //            'For example, enter <>mm to display a 5.0 millimeter radial dimension as "5.0mm." 
            //            'If you entered mm <>, the dimension would be displayed as "mm 5.0." 
            //            'Use the <> mechanism for angular dimensions. 

            //            .Dimjust = 0
            //            'Controls the horizontal positioning of dimension text. 
            //            '0 -  Positions the text above the dimension line and center-justifies it between the extension lines 
            //            '1 -  Positions the text next to the first extension line 
            //            '2 -  Positions the text next to the second extension line 
            //            '3 -  Positions the text above and aligned with the first extension line 
            //            '4 -  Positions the text above and aligned with the second extension line 

            //            .Dimadec = 0 'Controls the number of precision places displayed in angular dimensions. (0-8)
            //            .Dimalt = False 'Controls the display of alternate units in dimensions. Off - Disables alternate units
            //            .Dimaltd = 2 'Controls the number of decimal places in alternate units. If DIMALT is turned on, DIMALTD sets the number of digits displayed to the right of the decimal point in the alternate measurement
            //            .Dimaltf = 25.4 'Controls the multiplier for alternate units. If DIMALT is turned on, DIMALTF multiplies linear dimensions by a factor to produce a value in an alternate system of measurement. The initial value represents the number of millimeters in an inch.
            //            .Dimaltmzf = 100
            //            .Dimaltrnd = 0 'Rounds off the alternate dimension units. 
            //            .Dimalttd = 2 'Sets the number of decimal places for the tolerance values in the alternate units of a dimension. 
            //            .Dimalttz = 0 'Controls suppression of zeros in tolerance values. 
            //            .Dimaltu = 2 'Sets the units format for alternate units of all dimension substyles except Angular. (2 - Decimal)
            //            .Dimaltz = 0 'Controls the suppression of zeros for alternate unit dimension values. 
            //            .Dimapost = "" 'Specifies a text prefix or suffix (or both) to the alternate dimension measurement for all types of dimensions except angular. 
            //            'For instance, if the current units are Architectural, DIMALT is on, DIMALTF is 25.4 (the number of millimeters per inch), DIMALTD is 2, and DIMPOST is set to "mm," a distance of 10 units would be displayed as 10"[254.00mm]. 
            //            'To turn off an established prefix or suffix (or both), set it to a single period (.). 
            //            .Dimarcsym = 0 'Controls display of the arc symbol in an arc length dimension. (0- Places arc length symbols before the dimension text )
            //            '1 - Places arc length symbols above the dimension text 
            //            '2 -  Suppresses the display of arc length symbols 

            //            .Dimatfit = 3
            //            'Determines how dimension text and arrows are arranged when space is not sufficient to place both within the extension lines. 
            //            '0 -  Places both text and arrows outside extension lines 
            //            '1 -  Moves arrows first, then text
            //            '2 -  Moves text first, then arrows
            //            '3 -  Moves either text or arrows, whichever fits best 
            //            'A leader is added to moved dimension text when DIMTMOVE is set to 1. 


            //            .Dimaunit = 0 'Sets the units format for angular dimensions. (0 - Decimal degrees)
            //            .Dimazin = 0 'Suppresses zeros for angular dimensions. 


            //            .Dimsah = False
            //            'Controls the display of dimension line arrowhead blocks. 
            //            'Off - Use arrowhead blocks set by DIMBLK
            //            'On - Use arrowhead blocks set by DIMBLK1 and DIMBLK2

            //            .Dimblk = Arrowid
            //            'Sets the arrowhead block displayed at the ends of dimension lines or leader lines. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). Arrowhead block entries and the names used to select them in the New, Modify, and Override Dimension Style dialog boxes are shown below. You can also enter the names of user-defined arrowhead blocks. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            '"" - Closed(filled)
            //            '"_DOT" - dot
            //            '"_DOTSMALL" - dot small
            //            '"_DOTBLANK" - dot blank
            //            '"_ORIGIN" - origin indicator
            //            '"_ORIGIN2" - origin indicator 2
            //            '"_OPEN" - open
            //            '"_OPEN90" - Right(angle)
            //            '"_OPEN30" - open 30
            //            '"_CLOSED" - Closed
            //            '"_SMALL" - dot small blank
            //            '"_NONE" - none
            //            '"_OBLIQUE" - oblique
            //            '"_BOXFILLED" - box filled
            //            '"_BOXBLANK" - box
            //            '"_CLOSEDBLANK" - Closed(blank)
            //            '"_DATUMFILLED" - datum triangle filled
            //            '"_DATUMBLANK" - datum triangle
            //            '"_INTEGRAL" - integral
            //            '"_ARCHTICK" - architectural tick


            //            .Dimblk1 = Arrowid
            //            'Sets the arrowhead for the first end of the dimension line when DIMSAH is on. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowheads, see DIMBLK. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            .Dimblk2 = Arrowid
            //            'Sets the arrowhead for the second end of the dimension line when DIMSAH is on. 
            //            'To return to the default, closed-filled arrowhead display, enter a single period (.). For a list of arrowhead entries, see DIMBLK. 
            //            'Note Annotative blocks cannot be used as custom arrowheads for dimensions or leaders. 
            //            .Dimldrblk = Arrowid ' Specifies the arrow type for leaders. 

            //            .Dimcen = 0.09 'Controls drawing of circle or arc center marks and centerlines by the DIMCENTER, DIMDIAMETER, and DIMRADIUS commands. 
            //            .Dimclrd = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) ' Assigns colors to dimension lines, arrowheads, and dimension leader lines
            //            .Dimclre = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension extension lines.
            //            .Dimclrt = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256) 'Assigns colors to dimension text


            //            .Dimdle = 0 'Sets the distance the dimension line extends beyond the extension line when oblique strokes are drawn instead of arrowheads. 
            //            .Dimdli = 0.38 'Controls the spacing of the dimension lines in baseline dimensions. 
            //            'Each dimension line is offset from the previous one by this amount, if necessary, to avoid drawing over it. Changes made with DIMDLI are not applied to existing dimensions
            //            .Dimdsep = ".c"
            //            'Specifies a single-character decimal separator to use when creating dimensions whose unit format is decimal
            //            'When prompted, enter a single character at the Command prompt. If dimension units is set to Decimal, the DIMDSEP character is used instead of the default decimal point.
            //            'If DIMDSEP is set to NULL (default value, reset by entering a period), the decimal point is used as the dimension separator
            //            .Dimexe = 0.18 'Specifies how far to extend the extension line beyond the dimension line. 
            //            .Dimexo = 0.0625 'Specifies how far extension lines are offset from origin points. 
            //            'With fixed-length extension lines, this value determines the minimum offset. 
            //            .Dimfrac = 0 'Sets the fraction format when DIMLUNIT is set to 4 (Architectural) or 5 (Fractional).
            //            '0 - Horizontal stacking
            //            '1 - Diagonal stacking
            //            '2 - Not stacked (for example, 1/2)


            //            .Dimfxlen = 1
            //            .DimfxlenOn = False

            //            .Dimgap = 0.09 'Sets the distance around the dimension text when the dimension line breaks to accommodate dimension text.
            //            .Dimjogang = 0.785398163 'Determines the angle of the transverse segment of the dimension line in a jogged radius dimension. 



            //            .Dimlfac = 1
            //            'Sets a scale factor for linear dimension measurements. 
            //            'All linear dimension distances, including radii, diameters, and coordinates, are multiplied by DIMLFAC before being converted to dimension text. Positive values of DIMLFAC are applied to dimensions in both model space and paper space; negative values are applied to paper space only. 
            //            'DIMLFAC applies primarily to nonassociative dimensions (DIMASSOC set 0 or 1). For nonassociative dimensions in paper space, DIMLFAC must be set individually for each layout viewport to accommodate viewport scaling. 
            //            'DIMLFAC has no effect on angular dimensions, and is not applied to the values held in DIMRND, DIMTM, or DIMTP. 

            //            .Dimltex1 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the first extension line. 
            //            .Dimltex2 = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the second extension line. 
            //            .Dimltype = ThisDrawing.Database.ByBlockLinetype 'Sets the linetype of the dimension line.

            //            .Dimlunit = 2
            //            'Sets units for all dimension types except Angular. 
            //            '1 Scientific
            //            '2 Decimal
            //            '3 Engineering
            //            '4 Architectural (always displayed stacked)
            //            '5 Fractional (always displayed stacked)
            //            '6 Microsoft Windows Desktop (decimal format using Control Panel settings for decimal separator and number grouping symbols) 


            //            .Dimlwd = LineWeight.ByBlock
            //            'Assigns lineweight to dimension lines. 
            //            '-3 Default (the LWDEFAULT value) 
            //            '-2 BYBLOCK
            //            '-1 BYLAYER

            //            .Dimlwe = LineWeight.ByBlock
            //            'Assigns lineweight to extension  lines. 
            //            '-3 Default (the LWDEFAULT value) 
            //            '-2 BYBLOCK
            //            '-1 BYLAYER



            //            .Dimmzf = 100


            // .Dimscale = 1
            //            'Sets the overall scale factor applied to dimensioning variables that specify sizes, distances, or offsets. 
            //            'Also affects the leader objects with the LEADER command. 
            //            'Use MLEADERSCALE to scale multileader objects created with the MLEADER command. 
            //            '0.0 - A reasonable default value is computed based on the scaling between the current model space viewport and paper space. 
            //            'If you are in paper space or model space and not using the paper space feature, the scale factor is 1.0. 
            //            '>0 - A scale factor is computed that leads text sizes, arrowhead sizes, and other scaled distances to plot at their face values. 
            //            'DIMSCALE does not affect measured lengths, coordinates, or angles. 
            //            'Use DIMSCALE to control the overall scale of dimensions. However, if the current dimension style is annotative, 
            //            'DIMSCALE is automatically set to zero and the dimension scale is controlled by the CANNOSCALE system variable. DIMSCALE cannot be set to a non-zero value when using annotative dimensions. 

            //            .Dimtdec = 0
            //            'Sets the number of decimal places to display in tolerance values for the primary units in a dimension. 
            //            'This system variable has no effect unless DIMTOL is set to On. The default for DIMTOL is Off. 

            //            .Dimtfac = 1
            //            'Specifies a scale factor for the text height of fractions and tolerance values relative to the dimension text height, as set by DIMTXT. 
            //            'For example, if DIMTFAC is set to 1.0, the text height of fractions and tolerances is the same height as the dimension text. 
            //            'If DIMTFAC is set to 0.7500, the text height of fractions and tolerances is three-quarters the size of dimension text. 
            //            .Dimtfill = 1
            //            'Controls the background of dimension text. 
            //            '0 -  No Background
            //            '1 -  The background color of the drawing 
            //            '2 -  The background specified by DIMTFILLCLR
            //            .Dimtfillclr = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByBlock, 0)

            //            .Dimtix = False
            //            'Draws text between extension lines. 
            //            'Off -  Varies with the type of dimension. 
            //            '        For linear and angular dimensions, text is placed inside the extension lines if there is sufficient room. 
            //            '        For radius and diameter dimensions that don't fit inside the circle or arc, DIMTIX has no effect and always forces the text outside the circle or arc. 
            //            'On -  Draws dimension text between the extension lines even if it would ordinarily be placed outside those lines 

            //            .Dimsoxd = False
            //            'Suppresses arrowheads if not enough space is available inside the extension lines. 
            //            'Off -  Arrowheads are not suppressed
            //            'On -  Arrowheads are suppressed
            //            'If not enough space is available inside the extension lines and DIMTIX is on, setting DIMSOXD to On suppresses the arrowheads. If DIMTIX is off, DIMSOXD has no effect. 


            //            .Dimtm = 0
            //            'Sets the minimum (or lower) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            //            'DIMTM accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
            //            'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower, and a plus sign is added to the DIMTP value if it is positive. 
            //            'For DIMTM, the program uses the negative of the value you enter (adding a minus sign if you specify a positive number and a plus sign if you specify a negative number). 

            //            .Dimtmove = 0
            //            'Sets dimension text movement rules. 
            //            '0 -  Moves the dimension line with dimension text
            //            '1 -  Adds a leader when dimension text is moved
            //            '2 -  Allows text to be moved freely without a leader

            //            .Dimtp = 0
            //            'Sets the maximum (or upper) tolerance limit for dimension text when DIMTOL or DIMLIM is on. 
            //            'DIMTP accepts signed values. If DIMTOL is on and DIMTP and DIMTM are set to the same value, a tolerance value is drawn. 
            //            'If DIMTM and DIMTP values differ, the upper tolerance is drawn above the lower and a plus sign is added to the DIMTP value if it is positive. 


            //            .Dimlim = False
            //            'Generates dimension limits as the default text. 
            //            'Setting DIMLIM to On turns DIMTOL off. 
            //            'Off -  Dimension limits are not generated as default text 
            //            'On -  Dimension limits are generated as default text


            //            .Dimtol = False
            //            'Appends tolerances to dimension text. 
            //            'Setting DIMTOL to on turns DIMLIM off. 

            //            .Dimtolj = 1 'Sets the vertical justification for tolerance values relative to the nominal dimension text. 



            //            .Dimtsz = 0
            //            'Specifies the size of oblique strokes drawn instead of arrowheads for linear, radius, and diameter dimensioning. 
            //            '0 -  Draws arrowheads.
            //            '>0 -  Draws oblique strokes instead of arrowheads. The size of the oblique strokes is determined by this value multiplied by the DIMSCALE value 




            //            .Dimtzin = 0 'Controls the suppression of zeros in tolerance values. 

            //            .Dimupt = False
            //            'Controls options for user-positioned text. 
            //            'Off -  Cursor controls only the dimension line location
            //            'On -  Cursor controls both the text position and the dimension line location 

            //            .Dimzin = 0
            //            'Controls the suppression of zeros in the primary unit value. 
            //            'Values 0-3 affect feet-and-inch dimensions only: 
            //            '0 -  Suppresses zero feet and precisely zero inches
            //            '1 -  Includes zero feet and precisely zero inches
            //            '2 -  Includes zero feet and suppresses zero inches
            //            '3 -  Includes zero inches and suppresses zero feet
            //            '4 -  Suppresses leading zeros in decimal dimensions (for example, 0.5000 becomes .5000) 
            //            '8 -  Suppresses trailing zeros in decimal dimensions (for example, 12.5000 becomes 12.5) 
            //            '12 -  Suppresses both leading and trailing zeros (for example, 0.5000 becomes .5) 



            //        End With
        }

        public MLeader creaza_mleader(Point3d pt_ins, string continut, double texth, double delta_x, double delta_y, double lgap, double dogl, double arrow, string layer1)
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = new MLeader();


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                MText mtext1 = new MText();

                mtext1.Contents = continut;
                mtext1.TextHeight = texth;
                mtext1.BackgroundFill = true;
                mtext1.UseBackgroundColor = true;
                mtext1.BackgroundScaleFactor = 1.2;
                mtext1.ColorIndex = 0;

                mleader1.SetDatabaseDefaults();
                int index1 = mleader1.AddLeader();
                int index2 = mleader1.AddLeaderLine(index1);
                mleader1.AddFirstVertex(index2, pt_ins);
                mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + delta_x, pt_ins.Y + delta_y, pt_ins.Z));
                mleader1.LeaderLineType = LeaderType.StraightLeader;
                mleader1.ContentType = ContentType.MTextContent;
                mleader1.MText = mtext1;
                mleader1.TextHeight = texth;
                mleader1.LandingGap = lgap;
                mleader1.ArrowSize = arrow;
                mleader1.DoglegLength = dogl;
                mleader1.Annotative = AnnotativeStates.False;
                mleader1.ColorIndex = 256;
                mleader1.Layer = layer1;

                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);

                mleader1.MoveMLeader(new Vector3d(), MoveType.MoveAllPoints);
                mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.RightLeader);
                mleader1.SetTextAttachmentType(TextAttachmentType.AttachmentMiddleOfTop, LeaderDirectionType.LeftLeader);

                Trans1.Commit();

            }
            return mleader1;
        }

        public BlockReference InsertBlock_with_2scales(Database Database1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord, string NumeBlock, Point3d Insertion_point, double Scale_x, double Scale_y, double Rotation1, string Layer1)
        {

            BlockReference Block1 = null;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                if (BlockTable1.Has(NumeBlock) == true)
                {


                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTR = (BlockTableRecord)Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForRead);

                    Block1 = new BlockReference(Insertion_point, BTR.ObjectId);
                    Block1.Layer = Layer1;
                    Block1.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(Scale_x, Scale_y, 1);
                    Block1.Rotation = Rotation1;
                    BTrecord.AppendEntity(Block1);
                    Trans1.AddNewlyCreatedDBObject(Block1, true);

                }

                Trans1.Commit();
            }

            return Block1;
        }

        private void button_pipe_int_Click(object sender, EventArgs e)
        {

            if (refcl != null)
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

                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSelect Intersection with Pipe:");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status == PromptStatus.OK)
                            {
                                pipe_sta = refcl.GetDistAtPoint(refcl.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false));
                                textBox_cl_sta.Text = Functions.Get_chainage_from_double(pipe_sta, "f", 2);
                                label_int.Visible = true;
                            }
                            else
                            {
                                pipe_sta = -1.234;
                                textBox_cl_sta.Text = "";
                                label_int.Visible = false;
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
                set_enable_true();
            }
        }

        private void button_LOD_Click(object sender, EventArgs e)
        {

            if (refcl != null)
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


                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the lod polylines:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                            if (Rezultat1.Status == PromptStatus.OK)
                            {
                                for (int i = 0; i < Rezultat1.Value.Count; ++i)
                                {
                                    Polyline poly1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                    if (poly1 != null)
                                    {
                                        Point3dCollection col1 = Functions.Intersect_on_both_operands(refcl, poly1);
                                        if (col1.Count > 0)
                                        {
                                            if (lista_lod_sta == null) lista_lod_sta = new List<double>();
                                            for (int j = 0; j < col1.Count; ++j)
                                            {
                                                lista_lod_sta.Add(refcl.GetDistAtPoint(refcl.GetClosestPointTo(col1[j], Vector3d.ZAxis, false)));
                                            }

                                        }
                                    }
                                }
                            }
                            Trans1.Commit();
                            if (lista_lod_sta != null && lista_lod_sta.Count > 0)
                            {
                                label_LOD.Visible = true;

                            }
                            else
                            {
                                lista_lod_sta = null;
                                label_LOD.Visible = false;

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

        private void TextBox_keypress_only_integers(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_integer_pozitive_at_keypress(sender, e);
        }

        private void TextBox_keypress_only_doubles(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }

        private void button_stationing_Click(object sender, EventArgs e)
        {


            if (Functions.IsNumeric(textBox_spacing_major.Text) == false)
            {
                MessageBox.Show("spacing major issue");
                return;
            }


            double spacing_major = Math.Abs(Convert.ToDouble(textBox_spacing_major.Text));

            if (spacing_major == 0)
            {
                MessageBox.Show("spacing major issue");
                return;
            }

            if (Functions.IsNumeric(textBox_spacing_minor.Text) == false)
            {
                MessageBox.Show("spacing minor issue");
                return;
            }


            double spacing_minor = Math.Abs(Convert.ToDouble(textBox_spacing_minor.Text));

            if (spacing_minor == 0)
            {
                MessageBox.Show("spacing minor issue");
                return;
            }



            if (Functions.IsNumeric(textBox_tic_major.Text) == false)
            {
                MessageBox.Show("tick major issue");
                return;
            }


            double tick_major = Math.Abs(Convert.ToDouble(textBox_tic_major.Text));

            if (tick_major == 0)
            {
                MessageBox.Show("tick major issue");
                return;
            }

            if (Functions.IsNumeric(textBox_tic_minor.Text) == false)
            {
                MessageBox.Show("tick minor issue");
                return;
            }


            double tick_minor = Math.Abs(Convert.ToDouble(textBox_tic_minor.Text));

            if (tick_minor == 0)
            {
                MessageBox.Show("tick minor issue");
                return;
            }


            double texth = get_text_height();
            double gap1 = texth;





            double start1 = 0;


            set_enable_false();




            try
            {


                Functions.Creaza_layer(layer_stationing, 2, true);


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline 2d or 3d!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            Curve ent1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Curve;
                            Polyline poly1 = null;
                            if (ent1 is Polyline3d)
                            {
                                Polyline3d poly3 = ent1 as Polyline3d;
                                poly1 = Functions.Build_2dpoly_from_3d(poly3);
                            }

                            if (ent1 is Polyline)
                            {
                                Polyline poly3 = ent1 as Polyline;
                                poly1 = poly3;
                            }

                            if (poly1 != null)
                            {
                                create_stationing_2D(poly1, start1, gap1, texth, spacing_major, spacing_minor, tick_major, tick_minor);
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

            set_enable_true();

        }

        private void create_stationing_2D(Polyline Poly2D, double start1, double gap1, double texth, double spacing_major, double spacing_minor, double tick_major, double tick_minor)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);





                    double first_label_major = Math.Floor((start1 + spacing_major) / spacing_major) * spacing_major;

                    if (start1 + Poly2D.Length >= first_label_major)
                    {
                        int no_major = Convert.ToInt32(Math.Ceiling((start1 + Poly2D.Length - first_label_major) / spacing_major));

                        if (no_major > 0)
                        {
                            for (int i = 0; i < no_major; ++i)
                            {
                                Point3d pt0 = Poly2D.GetPointAtDist((first_label_major - start1) + i * spacing_major);


                                double label_major = first_label_major + i * spacing_major;
                                Line Big1 = new Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                double param2 = param1 + 1;
                                if (Poly2D.EndParam < param2)
                                {
                                    param1 = Poly2D.EndParam - 1;
                                    param2 = Poly2D.EndParam;
                                }

                                Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                double rot1 = bear1 - Math.PI / 2;

                                Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                Big1.Layer = layer_stationing;
                                Big1.ColorIndex = 256;


                                BTrecord.AppendEntity(Big1);
                                Trans1.AddNewlyCreatedDBObject(Big1, true);



                                Line l_t = new Line(Big1.StartPoint, Big1.EndPoint);
                                l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1);

                                mt1.Layer = layer_stationing;
                                BTrecord.AppendEntity(mt1);
                                Trans1.AddNewlyCreatedDBObject(mt1, true);



                            }
                        }
                    }

                    double first_label_minor = Math.Floor((start1 + spacing_minor) / spacing_minor) * spacing_minor;

                    if (start1 + Poly2D.Length >= first_label_minor)
                    {
                        int no_minor = Convert.ToInt32(Math.Ceiling((start1 + Poly2D.Length - first_label_minor) / spacing_minor));

                        if (no_minor > 0)
                        {
                            for (int i = 0; i < no_minor; ++i)
                            {
                                Point3d pt0 = Poly2D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                double label_major = first_label_minor + i * spacing_minor;
                                Line small1 = new Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                double param2 = param1 + 1;
                                if (Poly2D.EndParam < param2)
                                {
                                    param1 = Poly2D.EndParam - 1;
                                    param2 = Poly2D.EndParam;
                                }


                                Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                double rot1 = bear1 - Math.PI / 2;

                                small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                small1.Layer = layer_stationing;
                                small1.ColorIndex = 256;


                                BTrecord.AppendEntity(small1);
                                Trans1.AddNewlyCreatedDBObject(small1, true);



                            }
                        }
                    }


                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        public static MText creaza_mtext_sta(Point3d pt_ins, string continut, double texth, double rot1)
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

        private void button1_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
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
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.RotatedDimension), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            RotatedDimension dim1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as RotatedDimension;

                            if (dim1 != null)
                            {

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


        }

        private void button_stream_cross_section_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a 3D(2D) polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            dt_cross = null;
                            textBox_el_top2.Text = "";
                            textBox_el_bottom2.Text = "";
                            label_cross.Visible = false;
                            return;
                        }
                        Polyline3d p3 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;

                        if (p3 == null)
                        {
                            Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (p2 != null)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez_contours;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the contour polylines:";
                                Prompt_rez.SingleOnly = false;
                                rez_contours = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (rez_contours.Status != PromptStatus.OK)
                                {
                                    label_cross.Visible = false;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    dt_cross = null;
                                    textBox_el_top2.Text = "";
                                    textBox_el_bottom2.Text = "";
                                    set_enable_true();
                                    return;
                                }

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("sta", typeof(double));
                                dt2.Columns.Add("elev", typeof(double));
                                dt2.Columns.Add("pt", typeof(Point3d));
                                dt2.Columns.Add("dist", typeof(double));

                                for (int i = 0; i < rez_contours.Value.Count; ++i)
                                {
                                    #region polylines and elevations
                                    Polyline poly_cont = Trans1.GetObject(rez_contours.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                    if (poly_cont != null)
                                    {
                                        if (poly_cont.ObjectId != p2.ObjectId)
                                        {
                                            p2.Elevation = poly_cont.Elevation;

                                            Point3dCollection colint = Functions.Intersect_on_both_operands(p2, poly_cont);


                                            if (colint.Count > 0)
                                            {
                                                for (int j = 0; j < colint.Count; ++j)
                                                {
                                                    dt2.Rows.Add();
                                                    dt2.Rows[dt2.Rows.Count - 1]["sta"] = p2.GetDistAtPoint(colint[j]);
                                                    dt2.Rows[dt2.Rows.Count - 1]["elev"] = colint[j].Z;
                                                    dt2.Rows[dt2.Rows.Count - 1]["pt"] = colint[j];
                                                }


                                            }

                                            if (dt2.Rows.Count > 0) dt2 = Functions.Sort_data_table(dt2, "sta");

                                            if (dt2.Rows.Count > 0 && Math.Round(Convert.ToDouble(dt2.Rows[0]["sta"]), 3) > 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(1);
                                                Point3d pt0 = p2.GetPointAtParameter(0);

                                                Polyline poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_start.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_start.TransformBy(Matrix3d.Scaling(100 / poly_start.Length, pt0));
                                                poly_start.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_start, poly_cont);
                                                if (colint1.Count > 0)
                                                {
                                                    double dist_ex = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {

                                                        if (dt2.Rows[0]["dist"] != DBNull.Value)
                                                        {
                                                            dist_ex = Convert.ToDouble(dt2.Rows[0]["sta"]);
                                                        }
                                                        double dist1 = Math.Pow(Math.Pow(poly_start.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_start.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dt2.Rows[0]["dist"] == DBNull.Value)
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                System.Data.DataRow row1 = dt2.NewRow();
                                                                row1["sta"] = 0;
                                                                row1["pt"] = poly_start.StartPoint;

                                                                row1["dist"] = dist1;
                                                                dt2.Rows.InsertAt(row1, 0);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                dt2.Rows[0]["pt"] = poly_start.StartPoint;
                                                                dt2.Rows[0]["dist"] = dist1;
                                                            }
                                                        }


                                                    }

                                                }
                                            }

                                            if (dt2.Rows.Count == 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(1);
                                                Point3d pt0 = p2.GetPointAtParameter(0);

                                                Polyline poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_start.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_start.TransformBy(Matrix3d.Scaling(100 / poly_start.Length, pt0));
                                                poly_start.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_start, poly_cont);
                                                if (colint1.Count > 0)
                                                {
                                                    double dist2 = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {
                                                        double dist1 = Math.Pow(Math.Pow(poly_start.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_start.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dist1 < dist2)
                                                        {
                                                            dist2 = dist1;
                                                        }
                                                        System.Data.DataRow row1 = dt2.NewRow();
                                                        row1["sta"] = 0;
                                                        row1["pt"] = poly_start.StartPoint;
                                                        dt2.Rows.InsertAt(row1, 0);
                                                    }
                                                    dt2.Rows[0]["dist"] = dist2;
                                                }
                                            }

                                            if (dt2.Rows.Count > 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(p2.EndParam - 1);
                                                Point3d pt0 = p2.GetPointAtParameter(p2.EndParam);

                                                Polyline poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_end.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_end.TransformBy(Matrix3d.Scaling(100 / poly_end.Length, pt0));
                                                poly_end.Elevation = poly_cont.Elevation;


                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_end, poly_cont);

                                                if (colint1.Count > 0)
                                                {
                                                    double dist_ex = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {

                                                        if (dt2.Rows[dt2.Rows.Count - 1]["dist"] != DBNull.Value)
                                                        {
                                                            dist_ex = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["sta"]);
                                                        }
                                                        double dist1 = Math.Pow(Math.Pow(poly_end.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_end.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dt2.Rows[dt2.Rows.Count - 1]["dist"] == DBNull.Value)
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                System.Data.DataRow row1 = dt2.NewRow();
                                                                row1["sta"] = p2.Length;
                                                                row1["pt"] = poly_end.StartPoint;

                                                                row1["dist"] = dist1;
                                                                dt2.Rows.InsertAt(row1, dt2.Rows.Count);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                dt2.Rows[dt2.Rows.Count - 1]["pt"] = poly_end.StartPoint;
                                                                dt2.Rows[dt2.Rows.Count - 1]["dist"] = dist1;
                                                            }
                                                        }
                                                    }
                                                }
                                            }


                                            if (dt2.Rows.Count == 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(p2.EndParam - 1);
                                                Point3d pt0 = p2.GetPointAtParameter(p2.EndParam);

                                                Polyline poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_end.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_end.TransformBy(Matrix3d.Scaling(100 / poly_end.Length, pt0));
                                                poly_end.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_end, poly_cont);

                                                if (colint1.Count > 0)
                                                {
                                                    double dist2 = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {
                                                        double dist1 = Math.Pow(Math.Pow(poly_end.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_end.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dist1 < dist2)
                                                        {
                                                            dist2 = dist1;
                                                        }
                                                        System.Data.DataRow row1 = dt2.NewRow();
                                                        row1["sta"] = p2.Length;
                                                        row1["pt"] = poly_end.StartPoint;


                                                        dt2.Rows.InsertAt(row1, 0);
                                                    }
                                                    dt2.Rows[0]["dist"] = dist2;
                                                }
                                            }

                                        }
                                    }
                                    #endregion

                                    #region 3d polyline
                                    Polyline3d cont3D = Trans1.GetObject(rez_contours.Value[i].ObjectId, OpenMode.ForRead) as Polyline3d;
                                    if (cont3D != null)
                                    {
                                        Polyline cont2D = Functions.Build_2dpoly_from_3d(cont3D);
                                        cont2D.Elevation = p2.Elevation;
                                        Point3dCollection col_int3 = Functions.Intersect_on_both_operands(p2, cont2D);
                                        if (col_int3.Count > 0)
                                        {
                                            for (int j = 0; j < col_int3.Count; ++j)
                                            {
                                                double param1 = cont2D.GetParameterAtPoint(col_int3[j]);

                                                double z = cont3D.GetPointAtParameter(param1).Z;
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1]["sta"] = p2.GetDistAtPoint(col_int3[j]);
                                                dt2.Rows[dt2.Rows.Count - 1]["elev"] = z;
                                                dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(col_int3[j].X, col_int3[j].Y, z);

                                            }
                                        }

                                    }

                                    #endregion


                                }

                                if (dt2.Rows.Count > 1)
                                {


                                    dt2 = Functions.Sort_data_table(dt2, "sta");
                                    if (dt2.Rows[0]["dist"] != DBNull.Value)
                                    {
                                        double dist1 = Convert.ToDouble(dt2.Rows[0]["dist"]);
                                        Point3d pt1 = (Point3d)(dt2.Rows[0]["pt"]);
                                        double z1 = pt1.Z;
                                        double spacing = Convert.ToDouble(dt2.Rows[1]["sta"]) + dist1;
                                        Point3d pt2 = (Point3d)(dt2.Rows[1]["pt"]);
                                        double z2 = pt2.Z;
                                        dt2.Rows[0]["pt"] = new Point3d(pt1.X, pt1.Y, z1 + (dist1 * (z2 - z1)) / spacing);

                                    }

                                    if (dt2.Rows[dt2.Rows.Count - 1]["dist"] != DBNull.Value)
                                    {
                                        double dist1 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["dist"]);
                                        Point3d pt1 = (Point3d)(dt2.Rows[dt2.Rows.Count - 1]["pt"]);
                                        double z1 = pt1.Z;
                                        double spacing = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 2]["sta"]) + dist1;
                                        Point3d pt2 = (Point3d)(dt2.Rows[dt2.Rows.Count - 2]["pt"]);
                                        double z2 = pt2.Z;
                                        dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(pt1.X, pt1.Y, z1 + (dist1 * (z2 - z1)) / spacing);

                                    }

                                    if (p2.NumberOfVertices > 2 && dt2.Rows.Count > 2)
                                    {
                                        for (int i = 1; i < p2.NumberOfVertices - 1; ++i)
                                        {


                                            double sta_node = p2.GetDistanceAtParameter(i);
                                            double x = p2.GetPointAtParameter(i).X;
                                            double y = p2.GetPointAtParameter(i).Y;
                                            double z = -0.123;
                                            for (int j = 1; j < dt2.Rows.Count; ++j)
                                            {
                                                double sta1 = Convert.ToDouble(dt2.Rows[j - 1]["sta"]);
                                                double sta2 = Convert.ToDouble(dt2.Rows[j]["sta"]);

                                                if (sta_node > sta1 && sta_node < sta2)
                                                {
                                                    Point3d pt1 = (Point3d)(dt2.Rows[j - 1]["pt"]);
                                                    Point3d pt2 = (Point3d)(dt2.Rows[j]["pt"]);
                                                    double z1 = pt1.Z;
                                                    double z2 = pt2.Z;
                                                    z = z1 + ((sta_node - sta1) * (z2 - z1)) / (sta2 - sta1);
                                                    j = dt2.Rows.Count;
                                                }
                                            }

                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(x, y, z);
                                            dt2.Rows[dt2.Rows.Count - 1]["sta"] = sta_node;
                                        }

                                        dt2 = Functions.Sort_data_table(dt2, "sta");
                                    }



                                    p3 = new Polyline3d();
                                    p3.Layer = p2.Layer;
                                    BTrecord.AppendEntity(p3);
                                    Trans1.AddNewlyCreatedDBObject(p3, true);

                                    for (int i = 0; i < dt2.Rows.Count; ++i)
                                    {
                                        PolylineVertex3d Vertex_new = new PolylineVertex3d((Point3d)dt2.Rows[i]["pt"]);
                                        p3.AppendVertex(Vertex_new);
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                    }
                                    p2.UpgradeOpen();
                                    p2.Erase();
                                }
                            }
                        }


                        if (p3 != null)
                        {
                            Polyline crosscl = Build_2dpoly_from_3d(p3, textBox_el_bottom2, textBox_el_top2, 8, 2);
                            dt_cross = creaza_data_table(p3, crosscl, 2);
                            label_cross.Visible = true;
                        }
                        else
                        {
                            dt_cross = null;
                            label_cross.Visible = false;
                            textBox_el_top2.Text = "";
                            textBox_el_bottom2.Text = "";
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
            set_enable_true();
        }

        private void button_Eq_sect_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a 3D(2D) polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline3d), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            dt_eq = null;
                            textBox_el_top3.Text = "";
                            textBox_el_bottom3.Text = "";
                            label_eq.Visible = false;
                            return;
                        }
                        Polyline3d p3 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline3d;

                        if (p3 == null)
                        {
                            Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (p2 != null)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult rez_contours;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the contour polylines:";
                                Prompt_rez.SingleOnly = false;
                                rez_contours = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                if (rez_contours.Status != PromptStatus.OK)
                                {
                                    label_cross.Visible = false;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    dt_eq = null;
                                    textBox_el_top3.Text = "";
                                    textBox_el_bottom3.Text = "";
                                    label_eq.Visible = false;
                                    set_enable_true();
                                    return;
                                }

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("sta", typeof(double));
                                dt2.Columns.Add("elev", typeof(double));
                                dt2.Columns.Add("pt", typeof(Point3d));
                                dt2.Columns.Add("dist", typeof(double));

                                for (int i = 0; i < rez_contours.Value.Count; ++i)
                                {
                                    #region polylines and elevations
                                    Polyline poly_cont = Trans1.GetObject(rez_contours.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                    if (poly_cont != null)
                                    {
                                        if (poly_cont.ObjectId != p2.ObjectId)
                                        {
                                            p2.Elevation = poly_cont.Elevation;

                                            Point3dCollection colint = Functions.Intersect_on_both_operands(p2, poly_cont);


                                            if (colint.Count > 0)
                                            {
                                                for (int j = 0; j < colint.Count; ++j)
                                                {
                                                    dt2.Rows.Add();
                                                    dt2.Rows[dt2.Rows.Count - 1]["sta"] = p2.GetDistAtPoint(colint[j]);
                                                    dt2.Rows[dt2.Rows.Count - 1]["elev"] = colint[j].Z;
                                                    dt2.Rows[dt2.Rows.Count - 1]["pt"] = colint[j];
                                                }


                                            }

                                            if (dt2.Rows.Count > 0) dt2 = Functions.Sort_data_table(dt2, "sta");

                                            if (dt2.Rows.Count > 0 && Math.Round(Convert.ToDouble(dt2.Rows[0]["sta"]), 3) > 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(1);
                                                Point3d pt0 = p2.GetPointAtParameter(0);

                                                Polyline poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_start.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_start.TransformBy(Matrix3d.Scaling(100 / poly_start.Length, pt0));
                                                poly_start.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_start, poly_cont);
                                                if (colint1.Count > 0)
                                                {
                                                    double dist_ex = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {

                                                        if (dt2.Rows[0]["dist"] != DBNull.Value)
                                                        {
                                                            dist_ex = Convert.ToDouble(dt2.Rows[0]["sta"]);
                                                        }
                                                        double dist1 = Math.Pow(Math.Pow(poly_start.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_start.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dt2.Rows[0]["dist"] == DBNull.Value)
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                System.Data.DataRow row1 = dt2.NewRow();
                                                                row1["sta"] = 0;
                                                                row1["pt"] = poly_start.StartPoint;

                                                                row1["dist"] = dist1;
                                                                dt2.Rows.InsertAt(row1, 0);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                dt2.Rows[0]["pt"] = poly_start.StartPoint;
                                                                dt2.Rows[0]["dist"] = dist1;
                                                            }
                                                        }


                                                    }

                                                }
                                            }

                                            if (dt2.Rows.Count == 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(1);
                                                Point3d pt0 = p2.GetPointAtParameter(0);

                                                Polyline poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_start.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_start.TransformBy(Matrix3d.Scaling(100 / poly_start.Length, pt0));
                                                poly_start.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_start, poly_cont);
                                                if (colint1.Count > 0)
                                                {
                                                    double dist2 = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {
                                                        double dist1 = Math.Pow(Math.Pow(poly_start.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_start.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dist1 < dist2)
                                                        {
                                                            dist2 = dist1;
                                                        }
                                                        System.Data.DataRow row1 = dt2.NewRow();
                                                        row1["sta"] = 0;
                                                        row1["pt"] = poly_start.StartPoint;
                                                        dt2.Rows.InsertAt(row1, 0);
                                                    }
                                                    dt2.Rows[0]["dist"] = dist2;
                                                }
                                            }

                                            if (dt2.Rows.Count > 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(p2.EndParam - 1);
                                                Point3d pt0 = p2.GetPointAtParameter(p2.EndParam);

                                                Polyline poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_end.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_end.TransformBy(Matrix3d.Scaling(100 / poly_end.Length, pt0));
                                                poly_end.Elevation = poly_cont.Elevation;


                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_end, poly_cont);

                                                if (colint1.Count > 0)
                                                {
                                                    double dist_ex = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {

                                                        if (dt2.Rows[dt2.Rows.Count - 1]["dist"] != DBNull.Value)
                                                        {
                                                            dist_ex = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["sta"]);
                                                        }
                                                        double dist1 = Math.Pow(Math.Pow(poly_end.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_end.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dt2.Rows[dt2.Rows.Count - 1]["dist"] == DBNull.Value)
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                System.Data.DataRow row1 = dt2.NewRow();
                                                                row1["sta"] = p2.Length;
                                                                row1["pt"] = poly_end.StartPoint;

                                                                row1["dist"] = dist1;
                                                                dt2.Rows.InsertAt(row1, dt2.Rows.Count);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (dist1 < dist_ex)
                                                            {
                                                                dt2.Rows[dt2.Rows.Count - 1]["pt"] = poly_end.StartPoint;
                                                                dt2.Rows[dt2.Rows.Count - 1]["dist"] = dist1;
                                                            }
                                                        }
                                                    }
                                                }
                                            }


                                            if (dt2.Rows.Count == 0)
                                            {
                                                Point3d pt1 = p2.GetPointAtParameter(p2.EndParam - 1);
                                                Point3d pt0 = p2.GetPointAtParameter(p2.EndParam);

                                                Polyline poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt0.X, pt0.Y), 0, 0, 0);

                                                poly_end.TransformBy(Matrix3d.Displacement(pt1.GetVectorTo(pt0)));
                                                poly_end.TransformBy(Matrix3d.Scaling(100 / poly_end.Length, pt0));
                                                poly_end.Elevation = poly_cont.Elevation;

                                                Point3dCollection colint1 = Functions.Intersect_on_both_operands(poly_end, poly_cont);

                                                if (colint1.Count > 0)
                                                {
                                                    double dist2 = 1000;
                                                    for (int j = 0; j < colint1.Count; ++j)
                                                    {
                                                        double dist1 = Math.Pow(Math.Pow(poly_end.StartPoint.X - colint1[j].X, 2) + Math.Pow(poly_end.StartPoint.Y - colint1[j].Y, 2), 0.5);
                                                        if (dist1 < dist2)
                                                        {
                                                            dist2 = dist1;
                                                        }
                                                        System.Data.DataRow row1 = dt2.NewRow();
                                                        row1["sta"] = p2.Length;
                                                        row1["pt"] = poly_end.StartPoint;


                                                        dt2.Rows.InsertAt(row1, 0);
                                                    }
                                                    dt2.Rows[0]["dist"] = dist2;
                                                }
                                            }

                                        }
                                    }
                                    #endregion

                                    #region 3d polyline
                                    Polyline3d cont3D = Trans1.GetObject(rez_contours.Value[i].ObjectId, OpenMode.ForRead) as Polyline3d;
                                    if (cont3D != null)
                                    {
                                        Polyline cont2D = Functions.Build_2dpoly_from_3d(cont3D);
                                        cont2D.Elevation = p2.Elevation;
                                        Point3dCollection col_int3 = Functions.Intersect_on_both_operands(p2, cont2D);
                                        if (col_int3.Count > 0)
                                        {
                                            for (int j = 0; j < col_int3.Count; ++j)
                                            {
                                                double param1 = cont2D.GetParameterAtPoint(col_int3[j]);

                                                double z = cont3D.GetPointAtParameter(param1).Z;
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1]["sta"] = p2.GetDistAtPoint(col_int3[j]);
                                                dt2.Rows[dt2.Rows.Count - 1]["elev"] = z;
                                                dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(col_int3[j].X, col_int3[j].Y, z);

                                            }
                                        }

                                    }

                                    #endregion


                                }

                                if (dt2.Rows.Count > 1)
                                {


                                    dt2 = Functions.Sort_data_table(dt2, "sta");
                                    if (dt2.Rows[0]["dist"] != DBNull.Value)
                                    {
                                        double dist1 = Convert.ToDouble(dt2.Rows[0]["dist"]);
                                        Point3d pt1 = (Point3d)(dt2.Rows[0]["pt"]);
                                        double z1 = pt1.Z;
                                        double spacing = Convert.ToDouble(dt2.Rows[1]["sta"]) + dist1;
                                        Point3d pt2 = (Point3d)(dt2.Rows[1]["pt"]);
                                        double z2 = pt2.Z;
                                        dt2.Rows[0]["pt"] = new Point3d(pt1.X, pt1.Y, z1 + (dist1 * (z2 - z1)) / spacing);

                                    }

                                    if (dt2.Rows[dt2.Rows.Count - 1]["dist"] != DBNull.Value)
                                    {
                                        double dist1 = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 1]["dist"]);
                                        Point3d pt1 = (Point3d)(dt2.Rows[dt2.Rows.Count - 1]["pt"]);
                                        double z1 = pt1.Z;
                                        double spacing = Convert.ToDouble(dt2.Rows[dt2.Rows.Count - 2]["sta"]) + dist1;
                                        Point3d pt2 = (Point3d)(dt2.Rows[dt2.Rows.Count - 2]["pt"]);
                                        double z2 = pt2.Z;
                                        dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(pt1.X, pt1.Y, z1 + (dist1 * (z2 - z1)) / spacing);

                                    }

                                    if (p2.NumberOfVertices > 2 && dt2.Rows.Count > 2)
                                    {
                                        for (int i = 1; i < p2.NumberOfVertices - 1; ++i)
                                        {


                                            double sta_node = p2.GetDistanceAtParameter(i);
                                            double x = p2.GetPointAtParameter(i).X;
                                            double y = p2.GetPointAtParameter(i).Y;
                                            double z = -0.123;
                                            for (int j = 1; j < dt2.Rows.Count; ++j)
                                            {
                                                double sta1 = Convert.ToDouble(dt2.Rows[j - 1]["sta"]);
                                                double sta2 = Convert.ToDouble(dt2.Rows[j]["sta"]);

                                                if (sta_node > sta1 && sta_node < sta2)
                                                {
                                                    Point3d pt1 = (Point3d)(dt2.Rows[j - 1]["pt"]);
                                                    Point3d pt2 = (Point3d)(dt2.Rows[j]["pt"]);
                                                    double z1 = pt1.Z;
                                                    double z2 = pt2.Z;
                                                    z = z1 + ((sta_node - sta1) * (z2 - z1)) / (sta2 - sta1);
                                                    j = dt2.Rows.Count;
                                                }
                                            }

                                            dt2.Rows.Add();
                                            dt2.Rows[dt2.Rows.Count - 1]["pt"] = new Point3d(x, y, z);
                                            dt2.Rows[dt2.Rows.Count - 1]["sta"] = sta_node;
                                        }

                                        dt2 = Functions.Sort_data_table(dt2, "sta");
                                    }



                                    p3 = new Polyline3d();
                                    p3.Layer = p2.Layer;
                                    BTrecord.AppendEntity(p3);
                                    Trans1.AddNewlyCreatedDBObject(p3, true);

                                    for (int i = 0; i < dt2.Rows.Count; ++i)
                                    {
                                        PolylineVertex3d Vertex_new = new PolylineVertex3d((Point3d)dt2.Rows[i]["pt"]);
                                        p3.AppendVertex(Vertex_new);
                                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);
                                    }
                                    p2.UpgradeOpen();
                                    p2.Erase();
                                }
                            }
                        }


                        if (p3 != null)
                        {
                            Polyline crosscl = Build_2dpoly_from_3d(p3, textBox_el_bottom3, textBox_el_top3, 4, 2);
                            dt_eq = creaza_data_table(p3, crosscl, 3);
                            label_eq.Visible = true;
                        }
                        else
                        {
                            dt_eq = null;
                            label_eq.Visible = false;
                            textBox_el_top3.Text = "";
                            textBox_el_bottom3.Text = "";
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
            set_enable_true();
        }
    }
}
