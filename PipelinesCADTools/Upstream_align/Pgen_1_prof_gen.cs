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
    public partial class Pgen_prof_gen : Form
    {
        List<string> scales;
        Pgen_mainform Pg = null;

        System.Data.DataTable dt_pipe_cl;
        System.Data.DataTable dt_top_of_bank_ne;
        System.Data.DataTable dt_top_of_bank_sw;
        System.Data.DataTable dt_stream_cl;
        System.Data.DataTable dt_points;

        System.Data.DataTable dt_prof_stream_cl;
        System.Data.DataTable dt_prof_tob_ne;
        System.Data.DataTable dt_prof_tob_sw;

        System.Data.DataTable dt_prof_cont_stream_cl;
        System.Data.DataTable dt_prof_cont_tob_ne;
        System.Data.DataTable dt_prof_cont_tob_sw;

        double sta_cl = 1000;
        Point3d Ptint_stream_pipe;

        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(Button_Load_pipe_cl);
            lista_butoane.Add(button_load_TOB_en);
            lista_butoane.Add(button_load_TOB_ws);
            lista_butoane.Add(button_load_stream_cl);
            lista_butoane.Add(Button_load_survey_points);
            lista_butoane.Add(Button_draw_prof_streamcl);
            lista_butoane.Add(button_load_profile);
            lista_butoane.Add(button_load_contours);
            lista_butoane.Add(button_lod);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                if (sender as System.Windows.Forms.Button != bt1)
                {
                    bt1.Enabled = false;
                }
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(Button_Load_pipe_cl);
            lista_butoane.Add(button_load_TOB_en);
            lista_butoane.Add(button_load_TOB_ws);
            lista_butoane.Add(button_load_stream_cl);
            lista_butoane.Add(Button_load_survey_points);
            lista_butoane.Add(Button_draw_prof_streamcl);
            lista_butoane.Add(button_load_profile);
            lista_butoane.Add(button_load_contours);
            lista_butoane.Add(button_lod);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Pgen_prof_gen()
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
            Combobox_scales.SelectedIndex = 2;
        }

        private System.Data.DataTable creaza_data_table(Polyline poly1)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));
            if (poly1 != null)
            {
                for (int i = 0; i < poly1.NumberOfVertices; ++i)
                {
                    dt1.Rows.Add();
                    dt1.Rows[dt1.Rows.Count - 1][0] = poly1.GetPointAtParameter(i).X;
                    dt1.Rows[dt1.Rows.Count - 1][1] = poly1.GetPointAtParameter(i).Y;
                }
            }
            return dt1;
        }

        private void Button_Load_pipe_cl_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the pipeline centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            label_pipe.Visible = false;
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            dt_top_of_bank_ne = null;
                            set_enable_true();
                            return;
                        }
                        Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (p2 != null)
                        {
                            dt_pipe_cl = creaza_data_table(p2);
                        }
                        Trans1.Commit();
                        label_pipe.Visible = true;
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
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the top of bank [north//east]:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                        Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (p2 != null)
                        {
                            dt_top_of_bank_ne = creaza_data_table(p2);
                        }
                        Trans1.Commit();
                        label_tob_up.Visible = true;
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



        private void button_load_stream_cl_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect stream centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);
                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            label_stream.Visible = false;
                            dt_stream_cl = null;
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }
                        Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (p2 != null)
                        {
                            dt_stream_cl = creaza_data_table(p2);
                        }

                        if (dt_pipe_cl != null)
                        {
                            if (dt_pipe_cl.Rows.Count > 0)
                            {
                                Polyline pcl = Build_2d_poly_from_dt(dt_pipe_cl);
                                pcl.Elevation = p2.Elevation;
                                Point3dCollection colint = Functions.Intersect_on_both_operands(p2, pcl);
                                if (colint.Count > 0)
                                {
                                    Point3d pt1 = p2.GetClosestPointTo(colint[0], Vector3d.ZAxis, false);
                                    double Sta1 = p2.GetDistAtPoint(pt1);
                                    textBox_cl_sta.Text = Convert.ToString(Math.Round(Sta1, 3));
                                    Ptint_stream_pipe = colint[0];
                                }
                            }
                        }

                        Trans1.Commit();
                        label_stream.Visible = true;
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


        private void Button_load_survey_points_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Editor1.SetImpliedSelection(Empty_array);
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the survey information:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_points = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            return;
                        }
                        dt_points = new System.Data.DataTable();
                        dt_points.Columns.Add("x", typeof(double));
                        dt_points.Columns.Add("y", typeof(double));
                        dt_points.Columns.Add("z", typeof(double));
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (block1 != null)
                            {
                                if (block1.AttributeCollection.Count > 0)
                                {
                                    dt_points.Rows.Add();
                                    dt_points.Rows[dt_points.Rows.Count - 1][0] = block1.Position.X;
                                    dt_points.Rows[dt_points.Rows.Count - 1][1] = block1.Position.Y;
                                    foreach (ObjectId id1 in block1.AttributeCollection)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                        if (atr1 != null)
                                        {
                                            string tag1 = atr1.Tag;
                                            string val1 = atr1.TextString;
                                            if (tag1.ToLower() == "elev")
                                            {
                                                if (Functions.IsNumeric(val1) == true)
                                                {
                                                    dt_points.Rows[dt_points.Rows.Count - 1][2] = Convert.ToDouble(val1);
                                                }
                                            }
                                            else
                                            {
                                                if (tag1 != "")
                                                {
                                                    if (dt_points.Columns.Contains(tag1.ToUpper()) == false)
                                                    {
                                                        dt_points.Columns.Add(tag1.ToUpper(), typeof(string));
                                                    }
                                                    dt_points.Rows[dt_points.Rows.Count - 1][tag1.ToUpper()] = val1;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dtpoints);
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
            label_load_survey.Visible = true;

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

            if (Functions.IsNumeric(textBox_prof_Elev_top.Text) == false)
            {
                MessageBox.Show("please specify the top elevation");
                return;
            }
            if (Functions.IsNumeric(textBox_prof_Elev_bottom.Text) == false)
            {
                MessageBox.Show("please specify the bottom elevation");
                return;
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
                set_enable_false(sender);
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
                        string layer_ground = "_pgen_GROUND";
                        string layer_TOB_E_N = "_pgen_TOB_E_N";
                        string layer_TOB_W_S = "_pgen_TOB_W_S";
                        string layer_ground_cont = "_pgen_GROUND_cont";
                        string layer_TOB_E_N_cont = "_pgen_TOB_E_N_cont";
                        string layer_TOB_W_S_cont = "_pgen_TOB_W_S_cont";


                        double Downelev = Convert.ToDouble(textBox_prof_Elev_bottom.Text);
                        double Upelev = Convert.ToDouble(textBox_prof_Elev_top.Text);

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

                        draw_stream_cl_profile(
                            dt_prof_stream_cl, dt_prof_tob_ne, dt_prof_tob_sw,
                            dt_prof_cont_stream_cl, dt_prof_cont_tob_ne, dt_prof_cont_tob_sw,
                            pt_start, hincr, vincr,
                            Convert.ToDouble(textBox_prof_Hex.Text), Convert.ToDouble(textBox_prof_Vex.Text),
                            Downelev, Upelev,
                            layer_grid_lines, layer_text, layer_ground, layer_TOB_E_N, layer_TOB_W_S,
                            layer_ground_cont, layer_TOB_E_N_cont, layer_TOB_W_S_cont,
                            textH, Functions.Get_textstyle_id("Standard"), "", true, true, "f");



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

        private Polyline Build_2d_poly_from_dt(System.Data.DataTable dt_cl)
        {
            Polyline Poly2D = null;
            if (dt_cl != null && dt_cl.Rows.Count > 0)
            {
                Poly2D = new Polyline();
                int index1 = 0;

                for (int i = 0; i < dt_cl.Rows.Count; ++i)
                {
                    double x = 0;
                    double y = 0;

                    if (dt_cl.Rows[i][0] != DBNull.Value)
                    {
                        x = (double)dt_cl.Rows[i][0];
                        if (dt_cl.Rows[i][1] != DBNull.Value)
                        {
                            y = (double)dt_cl.Rows[i][1];

                            double bulge1 = 0;

                            Poly2D.AddVertexAt(index1, new Point2d(x, y), bulge1, 0, 0);


                            index1 = index1 + 1;
                        }
                    }
                }
                Poly2D.Elevation = 0;
            }

            return Poly2D;
        }

        private System.Data.DataTable creaza_dt_sta_and_elev_for_cl(Polyline poly1, Point3d pt0, double sta_label0)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("ptno", typeof(int));
            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));
            dt1.Columns.Add("sta", typeof(double));
            dt1.Columns.Add("elev", typeof(double));

            if (poly1.NumberOfVertices > 0)
            {
                if (dt_points != null)
                {
                    if (dt_points.Rows.Count > 0)
                    {

                        for (int i = 0; i < poly1.NumberOfVertices; ++i)
                        {
                            Point3d pt1 = poly1.GetPointAtParameter(i);
                            for (int j = 0; j < dt_points.Rows.Count; ++j)
                            {
                                if (dt_points.Rows[j][0] != DBNull.Value && dt_points.Rows[j][1] != DBNull.Value && dt_points.Rows[j][2] != DBNull.Value)
                                {
                                    double x = Convert.ToDouble(dt_points.Rows[j][0]);
                                    double y = Convert.ToDouble(dt_points.Rows[j][1]);
                                    double z = Convert.ToDouble(dt_points.Rows[j][2]);
                                    if (Math.Abs(pt1.X - x) < 0.1 && Math.Abs(pt1.Y - y) < 0.1)
                                    {
                                        dt1.Rows.Add();
                                        double sta1 = poly1.GetDistanceAtParameter(i);
                                        double sta0 = poly1.GetDistAtPoint(poly1.GetClosestPointTo(pt0, Vector3d.ZAxis, false));

                                        dt1.Rows[dt1.Rows.Count - 1][0] = dt1.Rows.Count;
                                        dt1.Rows[dt1.Rows.Count - 1][1] = poly1.GetPointAtParameter(i).X;
                                        dt1.Rows[dt1.Rows.Count - 1][2] = poly1.GetPointAtParameter(i).Y;

                                        dt1.Rows[dt1.Rows.Count - 1][3] = sta_label0 - (sta0 - sta1);
                                        dt1.Rows[dt1.Rows.Count - 1][4] = z;
                                    }


                                }
                            }
                        }

                        if (dt1.Rows.Count > 0)
                        {
                            #region don't used in the pgen - folow cl direction - no reverse
                            bool dont_use = true;
                            if (dont_use == false)
                            {
                                if (Convert.ToDouble(dt1.Rows[0][1]) < Convert.ToDouble(dt1.Rows[dt1.Rows.Count - 1][1]))
                                {
                                    //pt_start_prof = poly1.EndPoint;
                                    dt1.Rows.Clear();
                                    for (int i = 0; i < poly1.NumberOfVertices; ++i)
                                    {
                                        Point3d pt1 = poly1.GetPointAtParameter(i);
                                        for (int j = 0; j < dt_points.Rows.Count; ++j)
                                        {
                                            if (dt_points.Rows[j][0] != DBNull.Value && dt_points.Rows[j][1] != DBNull.Value && dt_points.Rows[j][2] != DBNull.Value)
                                            {
                                                double x = Convert.ToDouble(dt_points.Rows[j][0]);
                                                double y = Convert.ToDouble(dt_points.Rows[j][1]);
                                                double z = Convert.ToDouble(dt_points.Rows[j][2]);
                                                if (Math.Abs(pt1.X - x) < 0.1 && Math.Abs(pt1.Y - y) < 0.1)
                                                {
                                                    dt1.Rows.Add();
                                                    double sta1 = poly1.Length - poly1.GetDistanceAtParameter(i);
                                                    double sta0 = poly1.Length - poly1.GetDistAtPoint(poly1.GetClosestPointTo(pt0, Vector3d.ZAxis, false));

                                                    dt1.Rows[dt1.Rows.Count - 1][0] = poly1.NumberOfVertices - dt1.Rows.Count;
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = poly1.GetPointAtParameter(i).X;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = poly1.GetPointAtParameter(i).Y;

                                                    dt1.Rows[dt1.Rows.Count - 1][3] = sta_label0 - (sta0 - sta1);
                                                    dt1.Rows[dt1.Rows.Count - 1][4] = z;
                                                }


                                            }
                                        }
                                    }
                                }
                            }

                            #endregion

                            dt1 = Functions.Sort_data_table(dt1, "sta");
                        }
                    }
                }
            }


            return dt1;
        }


        private System.Data.DataTable creaza_dt_sta_and_elev_for_tob(Polyline poly_cl, Polyline poly_tob, Point3d pt0, double sta_label0)
        {
            System.Data.DataTable dt1 = null;

            if (poly_cl != null && poly_tob != null)
            {
                dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ptno", typeof(int));
                dt1.Columns.Add("x", typeof(double));
                dt1.Columns.Add("y", typeof(double));
                dt1.Columns.Add("sta", typeof(double));
                dt1.Columns.Add("elev", typeof(double));

                if (poly_cl.NumberOfVertices > 0 && poly_tob.NumberOfVertices > 0)
                {
                    if (dt_points != null)
                    {
                        if (dt_points.Rows.Count > 0)
                        {
                            for (int i = 0; i < poly_tob.NumberOfVertices; ++i)
                            {
                                Point3d pt1 = poly_tob.GetPointAtParameter(i);
                                for (int j = 0; j < dt_points.Rows.Count; ++j)
                                {
                                    if (dt_points.Rows[j][0] != DBNull.Value && dt_points.Rows[j][1] != DBNull.Value && dt_points.Rows[j][2] != DBNull.Value)
                                    {
                                        double x = Convert.ToDouble(dt_points.Rows[j][0]);
                                        double y = Convert.ToDouble(dt_points.Rows[j][1]);
                                        double z = Convert.ToDouble(dt_points.Rows[j][2]);
                                        if (Math.Abs(pt1.X - x) < 0.1 && Math.Abs(pt1.Y - y) < 0.1)
                                        {
                                            dt1.Rows.Add();

                                            Point3d pt_on_poly = poly_cl.GetClosestPointTo(poly_tob.GetPointAtParameter(i), Vector3d.ZAxis, false);

                                            double sta1 = poly_cl.GetDistAtPoint(pt_on_poly);
                                            double sta0 = poly_cl.GetDistAtPoint(poly_cl.GetClosestPointTo(pt0, Vector3d.ZAxis, false));

                                            dt1.Rows[dt1.Rows.Count - 1][0] = dt1.Rows.Count;
                                            dt1.Rows[dt1.Rows.Count - 1][1] = poly_tob.GetPointAtParameter(i).X;
                                            dt1.Rows[dt1.Rows.Count - 1][2] = poly_tob.GetPointAtParameter(i).Y;

                                            dt1.Rows[dt1.Rows.Count - 1][3] = sta_label0 - (sta0 - sta1);
                                            dt1.Rows[dt1.Rows.Count - 1][4] = z;
                                        }


                                    }
                                }
                            }
                        }
                    }
                }
            }




            return dt1;
        }

        static public void draw_stream_cl_profile(System.Data.DataTable dt_stream_surv, System.Data.DataTable dt_tob_ne_surv, System.Data.DataTable dt_tob_sw_surv,
                                                  System.Data.DataTable dt_stream_cont, System.Data.DataTable dt_tob_ne_cont, System.Data.DataTable dt_tob_sw_cont, Point3d Point0,
                                            double Hincr, double Vincr, double Hexag, double Vexag, double Downelev, double Upelev,
                                            string Layer_grid, string Layer_text,
                                            string Layer_poly_surv, string layer_tob_ne_surv, string layer_tob_sw_surv,
                                            string Layer_poly_cont, string layer_tob_ne_cont, string layer_tob_sw_cont,
                                            double Texth, ObjectId Textstyleid, string Elev_suffix,
                                            bool leftElev, bool rightElev, string units)
        {

            Functions.Creaza_layer(Layer_grid, 9, true);
            Functions.Creaza_layer(Layer_text, 2, true);

            if (dt_stream_surv != null && dt_stream_surv.Rows.Count > 1) Functions.Creaza_layer(Layer_poly_surv, 2, true);
            if (dt_tob_ne_surv != null && dt_tob_ne_surv.Rows.Count > 1) Functions.Creaza_layer(layer_tob_ne_surv, 3, true);
            if (dt_tob_sw_surv != null && dt_tob_sw_surv.Rows.Count > 1) Functions.Creaza_layer(layer_tob_sw_surv, 3, true);

            if (dt_stream_cont != null && dt_stream_cont.Rows.Count > 1) Functions.Creaza_layer(Layer_poly_cont, 2, true);
            if (dt_tob_ne_cont != null && dt_tob_ne_cont.Rows.Count > 1) Functions.Creaza_layer(layer_tob_ne_cont, 3, true);
            if (dt_tob_sw_cont != null && dt_tob_sw_cont.Rows.Count > 1) Functions.Creaza_layer(layer_tob_sw_cont, 3, true);

            double Startsta = 0;
            double Endsta = 0;
            double Textwidth = 0;

            double XR = Point0.X;

            string Col_sta = "sta";
            string Col_elev = "elev";

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                if ((dt_stream_surv != null && dt_stream_surv.Rows.Count > 0) || (dt_stream_cont != null && dt_stream_cont.Rows.Count > 0))
                {
                    double Min_sta = 0;
                    double Max_sta = 0;

                    if (dt_stream_surv != null && dt_stream_surv.Rows.Count > 0)
                    {
                        if (dt_stream_surv.Rows[0][Col_sta] != DBNull.Value)
                        {
                            Min_sta = Convert.ToDouble(dt_stream_surv.Rows[0][Col_sta]);
                        }

                        if (dt_stream_surv.Rows[dt_stream_surv.Rows.Count - 1][Col_sta] != DBNull.Value)
                        {
                            Max_sta = Convert.ToDouble(dt_stream_surv.Rows[dt_stream_surv.Rows.Count - 1][Col_sta]);
                        }
                    }

                    if (dt_stream_cont != null && dt_stream_cont.Rows.Count > 0)
                    {
                        if (dt_stream_cont.Rows[0][Col_sta] != DBNull.Value)
                        {
                            Min_sta = Convert.ToDouble(dt_stream_cont.Rows[0][Col_sta]);
                        }

                        if (dt_stream_cont.Rows[dt_stream_cont.Rows.Count - 1][Col_sta] != DBNull.Value)
                        {
                            Max_sta = Convert.ToDouble(dt_stream_cont.Rows[dt_stream_cont.Rows.Count - 1][Col_sta]);
                        }
                    }

                    Startsta = Functions.Round_Down_as_double(Min_sta, Hincr);
                    Endsta = Functions.Round_Up_as_double(Max_sta, Hincr);

                    int Nr_linii_elevation = Convert.ToInt32(((Upelev - Downelev) / Vincr) + 1);
                    int Nr_linii_station = Convert.ToInt32(((Endsta - Startsta) / Hincr) + 1);

                    double EndX = Point0.X + (Endsta - Startsta) * Hexag;

                    TextStyleTableRecord txtrec = Trans1.GetObject(Textstyleid, OpenMode.ForRead) as TextStyleTableRecord;



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
                        BTrecord.AppendEntity(LinieV);
                        Trans1.AddNewlyCreatedDBObject(LinieV, true);

                        MText Mt_sta = new MText();
                        Mt_sta.Contents = Functions.Get_chainage_from_double(DisplaySTA, units, 0);
                        Mt_sta.Layer = Layer_text;
                        Mt_sta.Attachment = AttachmentPoint.TopCenter;
                        Mt_sta.TextHeight = Texth;
                        Mt_sta.TextStyleId = Textstyleid;
                        Mt_sta.Location = new Point3d(Point0.X + PozX, Point0.Y - 2 * Texth, 0);
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
                    Polyline Poly_graph = new Polyline();
                    int idx_p = 0;


                    if (dt_stream_surv != null && dt_stream_surv.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_stream_surv.Rows.Count; ++i)
                        {
                            if (dt_stream_surv.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_stream_surv.Rows[i][Col_elev]);
                                if (dt_stream_surv.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_stream_surv.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph.Plinegen = true;
                        Poly_graph.Layer = Layer_poly_surv;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }



                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_tob_ne_surv != null && dt_tob_ne_surv.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_tob_ne_surv.Rows.Count; ++i)
                        {
                            if (dt_tob_ne_surv.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_tob_ne_surv.Rows[i][Col_elev]);
                                if (dt_tob_ne_surv.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_tob_ne_surv.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }

                        Poly_graph.Layer = layer_tob_ne_surv;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }

                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_tob_sw_surv != null && dt_tob_sw_surv.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_tob_sw_surv.Rows.Count; ++i)
                        {
                            if (dt_tob_sw_surv.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_tob_sw_surv.Rows[i][Col_elev]);
                                if (dt_tob_sw_surv.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_tob_sw_surv.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph.Plinegen = true;
                        Poly_graph.Layer = layer_tob_sw_surv;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }

                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_stream_cont != null && dt_stream_cont.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_stream_cont.Rows.Count; ++i)
                        {
                            if (dt_stream_cont.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_stream_cont.Rows[i][Col_elev]);
                                if (dt_stream_cont.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_stream_cont.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph.Plinegen = true;
                        Poly_graph.Layer = Layer_poly_cont;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }



                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_tob_ne_cont != null && dt_tob_ne_cont.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_tob_ne_cont.Rows.Count; ++i)
                        {
                            if (dt_tob_ne_cont.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_tob_ne_cont.Rows[i][Col_elev]);
                                if (dt_tob_ne_cont.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_tob_ne_cont.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph.Plinegen = true;
                        Poly_graph.Layer = layer_tob_ne_cont;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }

                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_tob_sw_cont != null && dt_tob_sw_cont.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_tob_sw_cont.Rows.Count; ++i)
                        {
                            if (dt_tob_sw_cont.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_tob_sw_cont.Rows[i][Col_elev]);
                                if (dt_tob_sw_cont.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_tob_sw_cont.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - Startsta) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }
                        Poly_graph.Plinegen = true;
                        Poly_graph.Layer = layer_tob_sw_cont;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }

                    #endregion



                }

                Trans1.Commit();
            }


        }

        private System.Data.DataTable Add_FSL_To_Matl_DB(System.Data.DataTable dt_fac0, System.Data.DataTable dt_mat0)
        {
            System.Data.DataTable dt_fac = dt_fac0.Copy();
            System.Data.DataTable dt_mat = dt_mat0.Copy();

            System.Data.DataTable dt3 = new System.Data.DataTable();
            dt3 = dt_mat.Clone();
            for (int i = dt_mat.Rows.Count - 1; i >= 0; --i)
            {
                double M_Start = Convert.ToDouble(dt_mat.Rows[i]["Begin Station"]);
                double M_End = Convert.ToDouble(dt_mat.Rows[i]["End Station"]);
                string descr = Convert.ToString(dt_mat.Rows[i]["Description"]);

                bool import = true;

                for (int j = 0; j < dt_fac.Rows.Count; ++j)
                {
                    double F_Start = Convert.ToDouble(dt_fac.Rows[j]["Begin"]);
                    double F_End = Convert.ToDouble(dt_fac.Rows[j]["End"]);



                    if ((M_Start < F_Start && M_End <= F_Start) || M_Start >= F_End)
                    {


                    }

                    else
                    {
                        import = false;
                        if (M_Start >= F_Start && M_End <= F_End)
                        {
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["Matl No"] = "REF";
                            dt_mat.Rows[i].Delete();
                        }
                        else if (M_Start < F_Start && M_End > F_Start && M_End <= F_End)
                        {
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["End Station"] = F_Start;
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["Begin Station"] = F_Start;
                            dt3.Rows[dt3.Rows.Count - 1]["Matl No"] = "REF";
                            dt_mat.Rows[i].Delete();
                        }
                        else if (M_Start >= F_Start && M_Start < F_End && M_End > F_End)
                        {
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["End Station"] = F_End;
                            dt3.Rows[dt3.Rows.Count - 1]["Matl No"] = "REF";
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["Begin Station"] = F_End;
                            dt_mat.Rows[i].Delete();
                        }
                        else if (M_Start < F_Start && M_End > F_End)
                        {
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["End Station"] = F_Start;
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["Begin Station"] = F_Start;
                            dt3.Rows[dt3.Rows.Count - 1]["End Station"] = F_End;
                            dt3.Rows[dt3.Rows.Count - 1]["Matl No"] = "REF";
                            dt3.ImportRow(dt_mat.Rows[i]);
                            dt3.Rows[dt3.Rows.Count - 1]["Begin Station"] = F_Start;
                            dt_mat.Rows[i].Delete();
                        }

                    }


                }
                if (import == true)
                {
                    dt3.ImportRow(dt_mat.Rows[i]);
                    dt_mat.Rows[i].Delete();
                }
            }


            return dt3;

        }

        private void button_load_profile_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Polyline poly_stream = Build_2d_poly_from_dt(dt_stream_cl);
                        Polyline poly_pipe = Build_2d_poly_from_dt(dt_pipe_cl);
                        Polyline poly_tob_up = Build_2d_poly_from_dt(dt_top_of_bank_ne);
                        Polyline poly_tob_down = Build_2d_poly_from_dt(dt_top_of_bank_sw);

                        string Col_elev = "elev";
                        double vincr = 1;

                        if (Functions.IsNumeric(textBox_cl_sta.Text) == true)
                        {
                            sta_cl = Convert.ToDouble(textBox_cl_sta.Text);
                        }

                        dt_prof_stream_cl = creaza_dt_sta_and_elev_for_cl(poly_stream, Ptint_stream_pipe, sta_cl);
                        dt_prof_tob_ne = creaza_dt_sta_and_elev_for_tob(poly_stream, poly_tob_up, Ptint_stream_pipe, sta_cl);
                        dt_prof_tob_sw = creaza_dt_sta_and_elev_for_tob(poly_stream, poly_tob_down, Ptint_stream_pipe, sta_cl);

                        double Downelev = 0;
                        double Upelev = 0;


                        if (dt_prof_cont_stream_cl != null)
                        {
                            if (dt_prof_cont_stream_cl.Rows.Count > 2)
                            {
                                double Min_el = 100000;
                                double Max_el = -100000;
                                for (int i = 0; i < dt_prof_cont_stream_cl.Rows.Count; ++i)
                                {
                                    if (dt_prof_cont_stream_cl.Rows[i][Col_elev] != DBNull.Value)
                                    {
                                        double z1 = Convert.ToDouble(dt_prof_cont_stream_cl.Rows[i][Col_elev]);
                                        if (z1 > Max_el) Max_el = z1;
                                        if (z1 < Min_el) Min_el = z1;
                                    }
                                }
                                Downelev = Functions.Round_Down_as_double(Min_el, vincr) - 8 * vincr;
                                Upelev = Functions.Round_Up_as_double(Max_el, vincr) + 5 * vincr;
                            }
                        }

                        else
                        {
                            if (dt_prof_stream_cl != null)
                            {
                                if (dt_prof_stream_cl.Rows.Count > 2)
                                {
                                    double Min_el = 100000;
                                    double Max_el = -100000;
                                    for (int i = 0; i < dt_prof_stream_cl.Rows.Count; ++i)
                                    {
                                        if (dt_prof_stream_cl.Rows[i][Col_elev] != DBNull.Value)
                                        {
                                            double z1 = Convert.ToDouble(dt_prof_stream_cl.Rows[i][Col_elev]);
                                            if (z1 > Max_el) Max_el = z1;
                                            if (z1 < Min_el) Min_el = z1;
                                        }
                                    }
                                    Downelev = Functions.Round_Down_as_double(Min_el, vincr) - 8 * vincr;
                                    Upelev = Functions.Round_Up_as_double(Max_el, vincr) + 5 * vincr;
                                }
                            }
                        }


                        textBox_prof_Elev_top.Text = Functions.Get_String_Rounded(Upelev, 0);
                        textBox_prof_Elev_bottom.Text = Functions.Get_String_Rounded(Downelev, 0);
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

        private void button_load_TOB_ws_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the top of bank [south//west]:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                        Polyline p2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (p2 != null)
                        {
                            dt_top_of_bank_sw = creaza_data_table(p2);
                        }
                        Trans1.Commit();
                        label_tob_down.Visible = true;
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

        private void button_load_contours_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            if (Functions.IsNumeric(textBox_cl_sta.Text) == true)
            {
                sta_cl = Convert.ToDouble(textBox_cl_sta.Text);
            }
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        Autodesk.Gis.Map.Project.ProjectModel project1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = project1.ODTables;

                        Editor1.SetImpliedSelection(Empty_array);
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the contours:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_points = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            return;
                        }


                        Polyline poly_pipe = Build_2d_poly_from_dt(dt_pipe_cl);

                        Polyline poly_stream = Build_2d_poly_from_dt(dt_stream_cl);
                        Polyline poly_tob_ne = Build_2d_poly_from_dt(dt_top_of_bank_ne);
                        Polyline poly_tob_sw = Build_2d_poly_from_dt(dt_top_of_bank_sw);


                        double sta0 = -1234.234;

                        if (poly_stream != null )
                        {
                            if (poly_pipe != null)
                            {
                                Point3dCollection col_o = Functions.Intersect_on_both_operands(poly_pipe, poly_stream);

                                if (col_o.Count > 0)
                                {
                                    Point3d Point0 = col_o[0];
                                    sta0 = poly_stream.GetDistAtPoint(poly_stream.GetClosestPointTo(Point0, Vector3d.ZAxis, false));
                                }
                            }
                            else
                            {
                                sta0 = 0;
                            }

                            #region poly contours

                            dt_prof_cont_stream_cl = new System.Data.DataTable();
                            dt_prof_cont_stream_cl.Columns.Add("ptno", typeof(int));
                            dt_prof_cont_stream_cl.Columns.Add("x", typeof(double));
                            dt_prof_cont_stream_cl.Columns.Add("y", typeof(double));
                            dt_prof_cont_stream_cl.Columns.Add("sta", typeof(double));
                            dt_prof_cont_stream_cl.Columns.Add("elev", typeof(double));

                            double sta_start_tob_ne = 0;
                            double sta_end_tob_ne = 0;

                            if (poly_tob_ne != null)
                            {
                                dt_prof_cont_tob_ne = new System.Data.DataTable();
                                dt_prof_cont_tob_ne.Columns.Add("ptno", typeof(int));
                                dt_prof_cont_tob_ne.Columns.Add("x", typeof(double));
                                dt_prof_cont_tob_ne.Columns.Add("y", typeof(double));
                                dt_prof_cont_tob_ne.Columns.Add("sta", typeof(double));
                                dt_prof_cont_tob_ne.Columns.Add("elev", typeof(double));

                                Point3d pt1_ne = poly_tob_ne.GetClosestPointTo(poly_stream.StartPoint, Vector3d.ZAxis, false);
                                Point3d pt2_ne = poly_tob_ne.GetClosestPointTo(poly_stream.EndPoint, Vector3d.ZAxis, false);

                                sta_start_tob_ne = poly_tob_ne.GetDistAtPoint(pt1_ne);
                                sta_end_tob_ne = poly_tob_ne.GetDistAtPoint(pt2_ne);

                                if (sta_end_tob_ne < sta_start_tob_ne)
                                {
                                    double t = sta_start_tob_ne;
                                    sta_start_tob_ne = sta_end_tob_ne;
                                    sta_end_tob_ne = t;
                                }


                            }

                            double sta_start_tob_sw = 0;
                            double sta_end_tob_sw = 0;

                            if (poly_tob_sw != null)
                            {
                                dt_prof_cont_tob_sw = new System.Data.DataTable();
                                dt_prof_cont_tob_sw.Columns.Add("ptno", typeof(int));
                                dt_prof_cont_tob_sw.Columns.Add("x", typeof(double));
                                dt_prof_cont_tob_sw.Columns.Add("y", typeof(double));
                                dt_prof_cont_tob_sw.Columns.Add("sta", typeof(double));
                                dt_prof_cont_tob_sw.Columns.Add("elev", typeof(double));

                                Point3d pt1_sw = poly_tob_sw.GetClosestPointTo(poly_stream.StartPoint, Vector3d.ZAxis, false);
                                Point3d pt2_sw = poly_tob_sw.GetClosestPointTo(poly_stream.EndPoint, Vector3d.ZAxis, false);

                                sta_start_tob_sw = poly_tob_sw.GetDistAtPoint(pt1_sw);
                                sta_end_tob_sw = poly_tob_sw.GetDistAtPoint(pt2_sw);

                                if (sta_end_tob_sw < sta_start_tob_sw)
                                {
                                    double t = sta_start_tob_sw;
                                    sta_start_tob_sw = sta_end_tob_sw;
                                    sta_end_tob_sw = t;
                                }


                            }



                            Polyline poly_start = new Polyline();
                            poly_start.AddVertexAt(0, poly_stream.GetPoint2dAt(0), 0, 0, 0);
                            poly_start.AddVertexAt(1, new Point2d(poly_stream.GetPoint2dAt(0).X + 1000, poly_stream.GetPoint2dAt(0).Y), 0, 0, 0);

                            double bear1 = Functions.GET_Bearing_rad(poly_stream.GetPoint2dAt(1).X, poly_stream.GetPoint2dAt(1).Y, poly_stream.GetPoint2dAt(0).X, poly_stream.GetPoint2dAt(0).Y);
                            poly_start.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly_start.StartPoint));

                            Polyline poly_end = new Polyline();
                            poly_end.AddVertexAt(0, poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 1), 0, 0, 0);
                            poly_end.AddVertexAt(1, new Point2d(poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 1).X + 1000, poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 1).Y), 0, 0, 0);

                            bear1 = Functions.GET_Bearing_rad(poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 2).X,
                                                                                           poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 2).Y,
                                                                                               poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 1).X,
                                                                                                   poly_stream.GetPoint2dAt(poly_stream.NumberOfVertices - 1).Y);

                            poly_end.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly_end.StartPoint));

                            Polyline poly_start_ne = null;
                            Polyline poly_end_ne = null;

                            if (poly_tob_ne != null)
                            {
                                poly_start_ne = new Polyline();
                                poly_end_ne = new Polyline();

                                poly_start_ne.AddVertexAt(0, poly_tob_ne.GetPoint2dAt(0), 0, 0, 0);
                                poly_start_ne.AddVertexAt(1, new Point2d(poly_tob_ne.GetPoint2dAt(0).X + 1000, poly_tob_ne.GetPoint2dAt(0).Y), 0, 0, 0);

                                double bear1_ne = Functions.GET_Bearing_rad(poly_tob_ne.GetPoint2dAt(1).X, poly_tob_ne.GetPoint2dAt(1).Y, poly_tob_ne.GetPoint2dAt(0).X, poly_tob_ne.GetPoint2dAt(0).Y);
                                poly_start_ne.TransformBy(Matrix3d.Rotation(bear1_ne, Vector3d.ZAxis, poly_start_ne.StartPoint));

                                poly_end_ne.AddVertexAt(0, poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 1), 0, 0, 0);
                                poly_end_ne.AddVertexAt(1, new Point2d(poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 1).X + 1000, poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 1).Y), 0, 0, 0);

                                bear1_ne = Functions.GET_Bearing_rad(poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 2).X,
                                                                                               poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 2).Y,
                                                                                                   poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 1).X,
                                                                                                       poly_tob_ne.GetPoint2dAt(poly_tob_ne.NumberOfVertices - 1).Y);

                                poly_end_ne.TransformBy(Matrix3d.Rotation(bear1_ne, Vector3d.ZAxis, poly_end_ne.StartPoint));
                            }

                            Polyline poly_start_sw = null;
                            Polyline poly_end_sw = null;

                            if (poly_tob_sw != null)
                            {
                                poly_start_sw = new Polyline();
                                poly_end_sw = new Polyline();

                                poly_start_sw.AddVertexAt(0, poly_tob_sw.GetPoint2dAt(0), 0, 0, 0);
                                poly_start_sw.AddVertexAt(1, new Point2d(poly_tob_sw.GetPoint2dAt(0).X + 1000, poly_tob_sw.GetPoint2dAt(0).Y), 0, 0, 0);

                                double bear1_sw = Functions.GET_Bearing_rad(poly_tob_sw.GetPoint2dAt(1).X, poly_tob_sw.GetPoint2dAt(1).Y, poly_tob_sw.GetPoint2dAt(0).X, poly_tob_sw.GetPoint2dAt(0).Y);
                                poly_start_sw.TransformBy(Matrix3d.Rotation(bear1_sw, Vector3d.ZAxis, poly_start_sw.StartPoint));

                                poly_end_sw.AddVertexAt(0, poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 1), 0, 0, 0);
                                poly_end_sw.AddVertexAt(1, new Point2d(poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 1).X + 1000, poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 1).Y), 0, 0, 0);

                                bear1_sw = Functions.GET_Bearing_rad(poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 2).X,
                                                                                               poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 2).Y,
                                                                                                   poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 1).X,
                                                                                                       poly_tob_sw.GetPoint2dAt(poly_tob_sw.NumberOfVertices - 1).Y);

                                poly_end_sw.TransformBy(Matrix3d.Rotation(bear1_sw, Vector3d.ZAxis, poly_end_sw.StartPoint));
                            }


                            double start_elev = -1234.234;
                            double calc_sta_start = 1234.234;
                            double end_elev = -1234.234;
                            double calc_sta_end = 1234.234;

                            double start_elev_ne = -1234.234;
                            double calc_sta_start_ne = 1234.234;
                            double end_elev_ne = -1234.234;
                            double calc_sta_end_ne = 1234.234;

                            double start_elev_sw = -1234.234;
                            double calc_sta_start_sw = 1234.234;
                            double end_elev_sw = -1234.234;
                            double calc_sta_end_sw = 1234.234;

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline poly_cont = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;

                                if (poly_cont != null)
                                {
                                    double elev1 = -1234.234;
                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat1.Value[i].ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
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
                                                        string Valoare1 = Record1[j].StrValue;
                                                        if (Nume_field.ToLower() == "elev")
                                                        {
                                                            if (Functions.IsNumeric(Valoare1) == true)
                                                            {
                                                                elev1 = Convert.ToDouble(Valoare1);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (elev1 == -1234.234)
                                    {
                                        elev1 = poly_cont.Elevation;
                                    }

                                    Polyline pcnt = new Polyline();
                                    pcnt = poly_cont.Clone() as Polyline;
                                    pcnt.Elevation = poly_stream.Elevation;

                                    Polyline pstart = new Polyline();
                                    pstart = poly_start.Clone() as Polyline;
                                    pstart.Elevation = pcnt.Elevation;

                                    Polyline pend = new Polyline();
                                    pend = poly_end.Clone() as Polyline;
                                    pend.Elevation = pcnt.Elevation;

                                    Polyline pcnt_ne = new Polyline();
                                    Polyline pstart_ne = new Polyline();
                                    Polyline pend_ne = new Polyline();

                                    if (poly_tob_ne != null)
                                    {
                                        pcnt_ne = poly_cont.Clone() as Polyline;
                                        pcnt_ne.Elevation = poly_tob_ne.Elevation;

                                        pstart_ne = poly_start_ne.Clone() as Polyline;
                                        pstart_ne.Elevation = pcnt_ne.Elevation;

                                        pend_ne = poly_end_ne.Clone() as Polyline;
                                        pend_ne.Elevation = pcnt_ne.Elevation;
                                    }


                                    Polyline pcnt_sw = new Polyline();
                                    Polyline pstart_sw = new Polyline();
                                    Polyline pend_sw = new Polyline();

                                    if (poly_tob_sw != null)
                                    {
                                        pcnt_sw = poly_cont.Clone() as Polyline;
                                        pcnt_sw.Elevation = poly_tob_sw.Elevation;

                                        pstart_sw = poly_start_sw.Clone() as Polyline;
                                        pstart_sw.Elevation = pcnt_sw.Elevation;

                                        pend_sw = poly_end_sw.Clone() as Polyline;
                                        pend_sw.Elevation = pcnt_sw.Elevation;
                                    }

                                    Point3dCollection colint = Functions.Intersect_on_both_operands(pcnt, poly_stream);
                                    if (colint.Count > 0)
                                    {
                                        for (int j = 0; j < colint.Count; ++j)
                                        {
                                            double sta1 = poly_stream.GetDistAtPoint(poly_stream.GetClosestPointTo(colint[j], Vector3d.ZAxis, false));
                                            dt_prof_cont_stream_cl.Rows.Add();
                                            dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["x"] = colint[j].X;
                                            dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["y"] = colint[j].Y;
                                            dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["elev"] = elev1;
                                            dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["ptno"] = Convert.ToInt32(sta1);
                                            dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["sta"] = sta_cl - (sta0 - sta1);
                                        }
                                    }

                                    Point3dCollection colint_start = Functions.Intersect_on_both_operands(pstart, pcnt);
                                    if (colint_start.Count > 0)
                                    {
                                        for (int j = 0; j < colint_start.Count; ++j)
                                        {
                                            double dist = pstart.GetDistAtPoint(pstart.GetClosestPointTo(colint_start[j], Vector3d.ZAxis, false));
                                            if (dist < calc_sta_start && dist > 0)
                                            {
                                                calc_sta_start = dist;
                                                start_elev = elev1;
                                            }
                                        }
                                    }

                                    Point3dCollection colint_end = Functions.Intersect_on_both_operands(pend, pcnt);
                                    if (colint_end.Count > 0)
                                    {
                                        for (int j = 0; j < colint_end.Count; ++j)
                                        {
                                            double dist = pend.GetDistAtPoint(pend.GetClosestPointTo(colint_end[j], Vector3d.ZAxis, false));
                                            if (dist < calc_sta_end && dist > 0)
                                            {
                                                calc_sta_end = dist;
                                                end_elev = elev1;
                                            }
                                        }
                                    }

                                    if (poly_tob_ne != null)
                                    {
                                        Point3dCollection colint_ne = Functions.Intersect_on_both_operands(pcnt_ne, poly_tob_ne);
                                        if (colint_ne.Count > 0)
                                        {
                                            for (int j = 0; j < colint_ne.Count; ++j)
                                            {
                                                double st0 = poly_tob_ne.GetDistAtPoint(poly_tob_ne.GetClosestPointTo(colint_ne[j], Vector3d.ZAxis, false));

                                                if (st0 >= sta_start_tob_ne && st0 <= sta_end_tob_ne)
                                                {
                                                    double sta1 = poly_stream.GetDistAtPoint(poly_stream.GetClosestPointTo(colint_ne[j], Vector3d.ZAxis, false));
                                                    dt_prof_cont_tob_ne.Rows.Add();
                                                    dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["x"] = colint_ne[j].X;
                                                    dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["y"] = colint_ne[j].Y;
                                                    dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["elev"] = elev1;
                                                    dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["ptno"] = Convert.ToInt32(sta1);
                                                    dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["sta"] = sta_cl - (sta0 - sta1);
                                                }
                                            }
                                        }

                                        Point3dCollection colint_start_ne = Functions.Intersect_on_both_operands(pstart_ne, pcnt_ne);
                                        if (colint_start_ne.Count > 0)
                                        {
                                            for (int j = 0; j < colint_start_ne.Count; ++j)
                                            {
                                                double dist = pstart_ne.GetDistAtPoint(pstart_ne.GetClosestPointTo(colint_start_ne[j], Vector3d.ZAxis, false));
                                                if (dist < calc_sta_start_ne && dist > 0)
                                                {
                                                    calc_sta_start_ne = dist;
                                                    start_elev_ne = elev1;
                                                }
                                            }
                                        }

                                        Point3dCollection colint_end_ne = Functions.Intersect_on_both_operands(pend_ne, pcnt_ne);
                                        if (colint_end_ne.Count > 0)
                                        {
                                            for (int j = 0; j < colint_end_ne.Count; ++j)
                                            {
                                                double dist = pend_ne.GetDistAtPoint(pend_ne.GetClosestPointTo(colint_end_ne[j], Vector3d.ZAxis, false));
                                                if (dist < calc_sta_end_ne && dist > 0)
                                                {
                                                    calc_sta_end_ne = dist;
                                                    end_elev_ne = elev1;
                                                }
                                            }
                                        }
                                    }

                                    if (poly_tob_sw != null)
                                    {
                                        Point3dCollection colint_sw = Functions.Intersect_on_both_operands(pcnt_sw, poly_tob_sw);
                                        if (colint_sw.Count > 0)
                                        {
                                            for (int j = 0; j < colint_sw.Count; ++j)
                                            {
                                                double st0 = poly_tob_sw.GetDistAtPoint(poly_tob_sw.GetClosestPointTo(colint_sw[j], Vector3d.ZAxis, false));
                                                if (st0 >= sta_start_tob_sw && st0 <= sta_end_tob_sw)
                                                {
                                                    double sta1 = poly_stream.GetDistAtPoint(poly_stream.GetClosestPointTo(colint_sw[j], Vector3d.ZAxis, false));
                                                    dt_prof_cont_tob_sw.Rows.Add();
                                                    dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["x"] = colint_sw[j].X;
                                                    dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["y"] = colint_sw[j].Y;
                                                    dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["elev"] = elev1;
                                                    dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["ptno"] = Convert.ToInt32(sta1);
                                                    dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["sta"] = sta_cl - (sta0 - sta1);
                                                }
                                            }
                                        }

                                        Point3dCollection colint_start_sw = Functions.Intersect_on_both_operands(pstart_sw, pcnt_sw);
                                        if (colint_start_sw.Count > 0)
                                        {
                                            for (int j = 0; j < colint_start_sw.Count; ++j)
                                            {
                                                double dist = pstart_sw.GetDistAtPoint(pstart_sw.GetClosestPointTo(colint_start_sw[j], Vector3d.ZAxis, false));
                                                if (dist < calc_sta_start_sw && dist > 0)
                                                {
                                                    calc_sta_start_sw = dist;
                                                    start_elev_sw = elev1;
                                                }
                                            }
                                        }

                                        Point3dCollection colint_end_sw = Functions.Intersect_on_both_operands(pend_sw, pcnt_sw);
                                        if (colint_end_sw.Count > 0)
                                        {
                                            for (int j = 0; j < colint_end_sw.Count; ++j)
                                            {
                                                double dist = pend_sw.GetDistAtPoint(pend_sw.GetClosestPointTo(colint_end_sw[j], Vector3d.ZAxis, false));
                                                if (dist < calc_sta_end_sw && dist > 0)
                                                {
                                                    calc_sta_end_sw = dist;
                                                    end_elev_sw = elev1;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (dt_prof_cont_stream_cl != null && dt_prof_cont_stream_cl.Rows.Count > 0)
                            {
                                dt_prof_cont_stream_cl = Functions.Sort_data_table(dt_prof_cont_stream_cl, "sta");
                                if (calc_sta_start < 1234.234)
                                {
                                    double elev1 = Convert.ToDouble(dt_prof_cont_stream_cl.Rows[0]["elev"]);
                                    double sta1 = Convert.ToDouble(dt_prof_cont_stream_cl.Rows[0]["sta"]);
                                    double stax = sta_cl - sta0;
                                    double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(start_elev - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_start);
                                    double elev2 = elev1 + delta_calc;
                                    if (elev1 > start_elev) elev2 = elev1 - delta_calc;

                                    System.Data.DataRow row0 = dt_prof_cont_stream_cl.NewRow();

                                    row0["x"] = poly_stream.StartPoint.X;
                                    row0["y"] = poly_stream.StartPoint.Y;
                                    row0["elev"] = elev2;
                                    row0["ptno"] = 0;
                                    row0["sta"] = stax;
                                    dt_prof_cont_stream_cl.Rows.InsertAt(row0, 0);
                                }
                                if (calc_sta_end < 1234.234)
                                {
                                    double elev1 = Convert.ToDouble(dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["elev"]);
                                    double sta1 = Convert.ToDouble(dt_prof_cont_stream_cl.Rows[dt_prof_cont_stream_cl.Rows.Count - 1]["sta"]);
                                    double stax = sta_cl - (sta0 - poly_stream.Length);
                                    double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(end_elev - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_end);
                                    double elev2 = elev1 + delta_calc;
                                    if (elev1 > end_elev) elev2 = elev1 - delta_calc;

                                    System.Data.DataRow row0 = dt_prof_cont_stream_cl.NewRow();
                                    row0["x"] = poly_stream.EndPoint.X;
                                    row0["y"] = poly_stream.EndPoint.Y;
                                    row0["elev"] = elev2;
                                    row0["ptno"] = Convert.ToInt32(poly_stream.Length);
                                    row0["sta"] = stax;
                                    dt_prof_cont_stream_cl.Rows.InsertAt(row0, dt_prof_cont_stream_cl.Rows.Count);
                                }
                            }

                            if (dt_prof_cont_tob_ne != null && dt_prof_cont_tob_ne.Rows.Count > 0)
                            {
                                dt_prof_cont_tob_ne = Functions.Sort_data_table(dt_prof_cont_tob_ne, "sta");
                                if (calc_sta_start_ne < 1234.234)
                                {
                                    if (Math.Round(sta_start_tob_ne, 3) == 0)
                                    {
                                        double elev1 = Convert.ToDouble(dt_prof_cont_tob_ne.Rows[0]["elev"]);
                                        double sta1 = Convert.ToDouble(dt_prof_cont_tob_ne.Rows[0]["sta"]);
                                        double stax = sta_cl - sta0;
                                        double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(start_elev_ne - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_start_ne);
                                        double elev2 = elev1 + delta_calc;
                                        if (elev1 > start_elev_ne) elev2 = elev1 - delta_calc;

                                        System.Data.DataRow row0 = dt_prof_cont_tob_ne.NewRow();
                                        row0["x"] = poly_tob_ne.StartPoint.X;
                                        row0["y"] = poly_tob_ne.StartPoint.Y;
                                        row0["elev"] = elev2;
                                        row0["ptno"] = 0;
                                        row0["sta"] = stax;
                                        dt_prof_cont_tob_ne.Rows.InsertAt(row0, 0);
                                    }
                                }

                                if (calc_sta_end_ne < 1234.234)
                                {
                                    if (Math.Round(sta_end_tob_ne, 3) == Math.Round(poly_tob_ne.Length, 3))
                                    {
                                        double elev1 = Convert.ToDouble(dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["elev"]);
                                        double sta1 = Convert.ToDouble(dt_prof_cont_tob_ne.Rows[dt_prof_cont_tob_ne.Rows.Count - 1]["sta"]);
                                        double stax = sta_cl - (sta0 - poly_stream.Length);
                                        double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(end_elev_ne - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_end_ne);
                                        double elev2 = elev1 + delta_calc;
                                        if (elev1 > end_elev_ne) elev2 = elev1 - delta_calc;

                                        System.Data.DataRow row0 = dt_prof_cont_tob_ne.NewRow();
                                        row0["x"] = poly_tob_ne.EndPoint.X;
                                        row0["y"] = poly_tob_ne.EndPoint.Y;
                                        row0["elev"] = elev2;
                                        row0["ptno"] = Convert.ToInt32(poly_tob_ne.Length);
                                        row0["sta"] = stax;
                                        dt_prof_cont_tob_ne.Rows.InsertAt(row0, dt_prof_cont_tob_ne.Rows.Count);
                                    }
                                }
                            }

                            if (dt_prof_cont_tob_sw != null && dt_prof_cont_tob_sw.Rows.Count > 0)
                            {
                                dt_prof_cont_tob_sw = Functions.Sort_data_table(dt_prof_cont_tob_sw, "sta");
                                if (calc_sta_start_sw < 1234.234)
                                {
                                    if (Math.Round(sta_start_tob_sw, 3) == 0)
                                    {
                                        double elev1 = Convert.ToDouble(dt_prof_cont_tob_sw.Rows[0]["elev"]);
                                        double sta1 = Convert.ToDouble(dt_prof_cont_tob_sw.Rows[0]["sta"]);
                                        double stax = sta_cl - sta0;
                                        double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(start_elev_sw - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_start_sw);
                                        double elev2 = elev1 + delta_calc;
                                        if (elev1 > start_elev_sw) elev2 = elev1 - delta_calc;

                                        System.Data.DataRow row0 = dt_prof_cont_tob_sw.NewRow();
                                        row0["x"] = poly_tob_sw.StartPoint.X;
                                        row0["y"] = poly_tob_sw.StartPoint.Y;
                                        row0["elev"] = elev2;
                                        row0["ptno"] = 0;
                                        row0["sta"] = stax;
                                        dt_prof_cont_tob_sw.Rows.InsertAt(row0, 0);
                                    }
                                }

                                if (calc_sta_end_sw < 1234.234)
                                {
                                    if (Math.Round(sta_end_tob_sw, 3) == Math.Round(poly_tob_sw.Length, 3))
                                    {
                                        double elev1 = Convert.ToDouble(dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["elev"]);
                                        double sta1 = Convert.ToDouble(dt_prof_cont_tob_sw.Rows[dt_prof_cont_tob_sw.Rows.Count - 1]["sta"]);
                                        double stax = sta_cl - (sta0 - poly_stream.Length);
                                        double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(end_elev_sw - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_end_sw);
                                        double elev2 = elev1 + delta_calc;
                                        if (elev1 > end_elev_sw) elev2 = elev1 - delta_calc;

                                        System.Data.DataRow row0 = dt_prof_cont_tob_sw.NewRow();
                                        row0["x"] = poly_tob_sw.EndPoint.X;
                                        row0["y"] = poly_tob_sw.EndPoint.Y;
                                        row0["elev"] = elev2;
                                        row0["ptno"] = Convert.ToInt32(poly_tob_sw.Length);
                                        row0["sta"] = stax;
                                        dt_prof_cont_tob_sw.Rows.InsertAt(row0, dt_prof_cont_tob_sw.Rows.Count);
                                    }
                                }
                            }
                        }
                        #endregion
                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_prof_cont_tob_ne);
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
            label_load_contours.Visible = true;
        }

        private void button_lod_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false(sender);
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        Autodesk.Gis.Map.Project.ProjectModel project1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = project1.ODTables;

                        Editor1.SetImpliedSelection(Empty_array);
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the limit of disturbance:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_points = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            return;
                        }





                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dtpoints);
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
            label_load_survey.Visible = true;

        }
    }
}
