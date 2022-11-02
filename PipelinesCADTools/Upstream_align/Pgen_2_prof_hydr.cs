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
    public partial class Pgen_prof_hyd : Form
    {
        List<string> scales;
        Pgen_mainform Pg = null;



        System.Data.DataTable dt_pts;
        System.Data.DataTable dt_prof_hydrant;
        System.Data.DataTable dt_prof_cont_ground;
        bool hydrant_is_pozitiv = true;

        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(Button_load_survey_points);
            lista_butoane.Add(Button_draw_prof_hydrant);
            lista_butoane.Add(button_load_profile);
            lista_butoane.Add(button_load_contours);
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
            lista_butoane.Add(Button_load_survey_points);
            lista_butoane.Add(Button_draw_prof_hydrant);
            lista_butoane.Add(button_load_profile);
            lista_butoane.Add(button_load_contours);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Pgen_prof_hyd()
        {
            InitializeComponent();
        }

        private void pgen_hydr_load(object sender, EventArgs e)
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
            Combobox_scales.SelectedIndex = 3;
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


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify first point on the centerline (Pipe direction)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_pts = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify second point on the centerline (Pipe direction)");
                        PP2.AllowNone = false;
                        PP2.BasePoint = Point_res1.Value;
                        PP2.UseBasePoint = true;
                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_pts = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_tee;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_tee = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_tee.MessageForAdding = "\nSelect the tee:";
                        Prompt_tee.SingleOnly = true;
                        Rezultat_tee = ThisDrawing.Editor.GetSelection(Prompt_tee);
                        if (Rezultat_tee.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_pts = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_valve;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_valve = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_valve.MessageForAdding = "\nSelect the valve:";
                        Prompt_valve.SingleOnly = true;
                        Rezultat_valve = ThisDrawing.Editor.GetSelection(Prompt_valve);
                        if (Rezultat_valve.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_pts = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_hyd;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_hyd = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_hyd.MessageForAdding = "\nSelect the hydrant:";
                        Prompt_hyd.SingleOnly = true;
                        Rezultat_hyd = ThisDrawing.Editor.GetSelection(Prompt_hyd);
                        if (Rezultat_hyd.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            dt_pts = null;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);
                            label_load_survey.Visible = false;
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            return;
                        }



                        dt_pts = new System.Data.DataTable();
                        dt_pts.Columns.Add("x", typeof(double));
                        dt_pts.Columns.Add("y", typeof(double));
                        dt_pts.Columns.Add("z", typeof(double));
                        dt_pts.Columns.Add("desc", typeof(string));

                        BlockReference block1 = Trans1.GetObject(Rezultat_hyd.Value[0].ObjectId, OpenMode.ForRead) as BlockReference;
                        DBPoint pt1 = Trans1.GetObject(Rezultat_hyd.Value[0].ObjectId, OpenMode.ForRead) as DBPoint;

                        if (block1 != null)
                        {

                            Polyline Poly_dir = new Polyline();
                            Poly_dir.AddVertexAt(0, new Point2d(Point_res1.Value.X, Point_res1.Value.Y), 0, 0, 0);
                            Poly_dir.AddVertexAt(1, new Point2d(Point_res2.Value.X, Point_res2.Value.Y), 0, 0, 0);



                            if (Functions.Angle_left_right(Poly_dir, block1.Position) == "LT.")
                            {
                                hydrant_is_pozitiv = false;
                            }
                            else
                            {
                                hydrant_is_pozitiv = true;
                            }

                            if (block1.AttributeCollection.Count > 0)
                            {
                                dt_pts.Rows.Add();
                                dt_pts.Rows[dt_pts.Rows.Count - 1][0] = block1.Position.X;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][1] = block1.Position.Y;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "hydrant";
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
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][2] = Convert.ToDouble(val1);
                                            }
                                        }
                                        else
                                        {
                                            if (tag1 != "")
                                            {
                                                if (dt_pts.Columns.Contains(tag1.ToUpper()) == false)
                                                {
                                                    dt_pts.Columns.Add(tag1.ToUpper(), typeof(string));
                                                }
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][tag1.ToUpper()] = val1;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (pt1 != null)
                        {
                            dt_pts.Rows.Add();
                            dt_pts.Rows[dt_pts.Rows.Count - 1][0] = pt1.Position.X;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][1] = pt1.Position.Y;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][2] = pt1.Position.Z;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "hydrant";
                        }

                        BlockReference block2 = Trans1.GetObject(Rezultat_valve.Value[0].ObjectId, OpenMode.ForRead) as BlockReference;
                        DBPoint pt2 = Trans1.GetObject(Rezultat_valve.Value[0].ObjectId, OpenMode.ForRead) as DBPoint;

                        if (block2 != null)
                        {
                            if (block2.AttributeCollection.Count > 0)
                            {
                                dt_pts.Rows.Add();
                                dt_pts.Rows[dt_pts.Rows.Count - 1][0] = block2.Position.X;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][1] = block2.Position.Y;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "valve";
                                foreach (ObjectId id1 in block2.AttributeCollection)
                                {
                                    AttributeReference atr2 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                    if (atr2 != null)
                                    {
                                        string tag2 = atr2.Tag;
                                        string val2 = atr2.TextString;
                                        if (tag2.ToLower() == "elev")
                                        {
                                            if (Functions.IsNumeric(val2) == true)
                                            {
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][2] = Convert.ToDouble(val2);
                                            }
                                        }
                                        else
                                        {
                                            if (tag2 != "")
                                            {
                                                if (dt_pts.Columns.Contains(tag2.ToUpper()) == false)
                                                {
                                                    dt_pts.Columns.Add(tag2.ToUpper(), typeof(string));
                                                }
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][tag2.ToUpper()] = val2;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (pt2 != null)
                        {
                            dt_pts.Rows.Add();
                            dt_pts.Rows[dt_pts.Rows.Count - 1][0] = pt2.Position.X;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][1] = pt2.Position.Y;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][2] = pt2.Position.Z;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "valve";
                        }

                        BlockReference block3 = Trans1.GetObject(Rezultat_tee.Value[0].ObjectId, OpenMode.ForRead) as BlockReference;
                        DBPoint pt3 = Trans1.GetObject(Rezultat_tee.Value[0].ObjectId, OpenMode.ForRead) as DBPoint;

                        if (block3 != null)
                        {
                            if (block3.AttributeCollection.Count > 0)
                            {
                                dt_pts.Rows.Add();
                                dt_pts.Rows[dt_pts.Rows.Count - 1][0] = block3.Position.X;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][1] = block3.Position.Y;
                                dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "tee";
                                foreach (ObjectId id1 in block3.AttributeCollection)
                                {
                                    AttributeReference atr3 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                    if (atr3 != null)
                                    {
                                        string tag3 = atr3.Tag;
                                        string val3 = atr3.TextString;
                                        if (tag3.ToLower() == "elev")
                                        {
                                            if (Functions.IsNumeric(val3) == true)
                                            {
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][2] = Convert.ToDouble(val3);
                                            }
                                        }
                                        else
                                        {
                                            if (tag3 != "")
                                            {
                                                if (dt_pts.Columns.Contains(tag3.ToUpper()) == false)
                                                {
                                                    dt_pts.Columns.Add(tag3.ToUpper(), typeof(string));
                                                }
                                                dt_pts.Rows[dt_pts.Rows.Count - 1][tag3.ToUpper()] = val3;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (pt3!=null)
                        {
                            dt_pts.Rows.Add();
                            dt_pts.Rows[dt_pts.Rows.Count - 1][0] = pt3.Position.X;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][1] = pt3.Position.Y;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][2] = pt3.Position.Z;
                            dt_pts.Rows[dt_pts.Rows.Count - 1][3] = "tee";
                        }

                  // Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_pts);
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


        private void Button_draw_prof_hydrant_Click(object sender, EventArgs e)
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
            if (Functions.IsNumeric(textBox_start_grid.Text) == false)
            {
                MessageBox.Show("please specify the start station for grid");
                return;
            }
            if (Functions.IsNumeric(textBox_end_grid.Text) == false)
            {
                MessageBox.Show("please specify the end station for grid");
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
                        string layer_grid_lines = "GRID_PROF_LINE";
                        string layer_grid_small_lines = "GRID_PROF_TICK";
                        string layer_grid_middle_lines = "G_PROF_GRID";
                        string layer_text = "GRID_PROF_LINE";
                        string layer_ground = "GROUND";
                        string layer_hydrant = "HYDRANT";
                        double sta1 = Convert.ToDouble(textBox_start_grid.Text);
                        double sta2 = Convert.ToDouble(textBox_end_grid.Text);


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

                        draw_hydrant_profile(
                            dt_prof_cont_ground, dt_prof_hydrant,
                            pt_start, hincr, vincr,
                            Convert.ToDouble(textBox_prof_Hex.Text), Convert.ToDouble(textBox_prof_Vex.Text),
                            Downelev, Upelev, sta1, sta2,
                            layer_grid_lines, layer_text, layer_ground, layer_hydrant, layer_grid_small_lines, layer_grid_middle_lines,
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

        private System.Data.DataTable creaza_dt_sta_and_elev_for_hydrant(Polyline poly1)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("ptno", typeof(int));
            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));
            dt1.Columns.Add("sta", typeof(double));
            dt1.Columns.Add("elev", typeof(double));

            if (poly1.NumberOfVertices > 0)
            {
                if (dt_pts != null)
                {
                    if (dt_pts.Rows.Count > 0)
                    {

                        for (int i = 0; i < poly1.NumberOfVertices; ++i)
                        {
                            Point3d pt1 = poly1.GetPointAtParameter(i);
                            double sta1 = poly1.GetDistanceAtParameter(i);
                            if (hydrant_is_pozitiv == false)
                            {
                                sta1 = -sta1;
                            }

                            for (int j = 0; j < dt_pts.Rows.Count; ++j)
                            {
                                if (dt_pts.Rows[j][0] != DBNull.Value && dt_pts.Rows[j][1] != DBNull.Value && dt_pts.Rows[j][2] != DBNull.Value)
                                {
                                    double x = Convert.ToDouble(dt_pts.Rows[j][0]);
                                    double y = Convert.ToDouble(dt_pts.Rows[j][1]);
                                    double z = Convert.ToDouble(dt_pts.Rows[j][2]);
                                    if (Math.Abs(pt1.X - x) < 0.1 && Math.Abs(pt1.Y - y) < 0.1)
                                    {
                                        dt1.Rows.Add();



                                        dt1.Rows[dt1.Rows.Count - 1][0] = dt1.Rows.Count;
                                        dt1.Rows[dt1.Rows.Count - 1][1] = pt1.X;
                                        dt1.Rows[dt1.Rows.Count - 1][2] = pt1.Y;

                                        dt1.Rows[dt1.Rows.Count - 1][3] = sta1;
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
                                        for (int j = 0; j < dt_pts.Rows.Count; ++j)
                                        {
                                            if (dt_pts.Rows[j][0] != DBNull.Value && dt_pts.Rows[j][1] != DBNull.Value && dt_pts.Rows[j][2] != DBNull.Value)
                                            {
                                                double x = Convert.ToDouble(dt_pts.Rows[j][0]);
                                                double y = Convert.ToDouble(dt_pts.Rows[j][1]);
                                                double z = Convert.ToDouble(dt_pts.Rows[j][2]);
                                                if (Math.Abs(pt1.X - x) < 0.1 && Math.Abs(pt1.Y - y) < 0.1)
                                                {
                                                    dt1.Rows.Add();
                                                    double sta1 = poly1.Length - poly1.GetDistanceAtParameter(i);


                                                    dt1.Rows[dt1.Rows.Count - 1][0] = poly1.NumberOfVertices - dt1.Rows.Count;
                                                    dt1.Rows[dt1.Rows.Count - 1][1] = poly1.GetPointAtParameter(i).X;
                                                    dt1.Rows[dt1.Rows.Count - 1][2] = poly1.GetPointAtParameter(i).Y;

                                                    dt1.Rows[dt1.Rows.Count - 1][3] = sta1;
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




        static public void draw_hydrant_profile(System.Data.DataTable dt_ground, System.Data.DataTable dt_hydrant, Point3d Point0,
                                            double Hincr, double Vincr, double Hexag, double Vexag, double Downelev, double Upelev,
                                            double start_grid, double end_grid,
                                            string Layer_grid, string Layer_text,
                                            string Layer_ground, string layer_hydrant, string layer_grid_small_lines, string layer_grid_middle_lines,
                                            double Texth, ObjectId Textstyleid, string Elev_suffix,
                                            bool leftElev, bool rightElev, string units)
        {

            Functions.Creaza_layer(Layer_grid, 27, true);
            Functions.Creaza_layer(Layer_text, 27, true);
            Functions.Creaza_layer(layer_grid_small_lines, 22, true);
            Functions.Creaza_layer(layer_grid_middle_lines, 13, true);

            if (dt_ground != null && dt_ground.Rows.Count > 1) Functions.Creaza_layer(Layer_ground, 3, true);
            if (dt_hydrant != null && dt_hydrant.Rows.Count > 1) Functions.Creaza_layer(layer_hydrant, 2, true);






            string Col_sta = "sta";
            string Col_elev = "elev";

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                if ((dt_ground != null && dt_ground.Rows.Count > 0) && (dt_hydrant != null && dt_hydrant.Rows.Count > 0))
                {




                    int Nr_linii_elevation = Convert.ToInt32(((Upelev - Downelev) / Vincr) + 1);
                    int Nr_linii_station = Convert.ToInt32(((end_grid - start_grid) / Hincr) + 1);

                    double EndX = Point0.X + (end_grid - start_grid) * Hexag;

                    TextStyleTableRecord txtrec = Trans1.GetObject(Textstyleid, OpenMode.ForRead) as TextStyleTableRecord;



                    #region station lines

                    for (int i = 0; i < Nr_linii_station; ++i)
                    {

                        double DisplaySTA = start_grid + i * Hincr;
                        double PozX = i * Hincr * Hexag;


                        Autodesk.AutoCAD.DatabaseServices.Line LinieV = new Autodesk.AutoCAD.DatabaseServices.Line(
                                                                                          new Point3d(Point0.X + PozX, Point0.Y, 0),
                                                                                          new Point3d(Point0.X + PozX, Point0.Y + (Upelev - Downelev) * Vexag, 0));

                        if (i == 0 || i == Nr_linii_station - 1)
                        {
                            LinieV.Layer = Layer_grid;
                        }
                        else
                        {
                            LinieV.Layer = layer_grid_middle_lines;
                            LinieV.LinetypeScale = 0.25;
                        }


                        LinieV.Linetype = "ByLayer";
                        BTrecord.AppendEntity(LinieV);
                        Trans1.AddNewlyCreatedDBObject(LinieV, true);

                        MText Mt_sta = new MText();
                        Mt_sta.Contents = Functions.Get_chainage_from_double(DisplaySTA, units, 0);
                        Mt_sta.Layer = Layer_text;
                        Mt_sta.Attachment = AttachmentPoint.TopCenter;
                        Mt_sta.TextHeight = Texth;
                        Mt_sta.TextStyleId = Textstyleid;
                        Mt_sta.Location = new Point3d(Point0.X + PozX, Point0.Y - 0.75 * Texth, 0);
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

                        if (i == 0)
                        {
                            LinieH.Layer = Layer_grid;
                        }
                        else
                        {
                            LinieH.Layer = layer_grid_middle_lines;
                            LinieH.LinetypeScale = 0.25;
                        }

                        LinieH.Linetype = "ByLayer";
                        BTrecord.AppendEntity(LinieH);
                        Trans1.AddNewlyCreatedDBObject(LinieH, true);

                        double scale1 = Texth / 0.08;
                        double y = Point0.Y + i * Vincr * Vexag;

                        if (i < Nr_linii_elevation - 1)
                        {


                            for (int j = 1; j <= 4; ++j)
                            {

                                Autodesk.AutoCAD.DatabaseServices.Line linie_small_left =
                                   new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(Point0.X, y + j * Vincr * Vexag / 5, 0),
                                                                              new Point3d(Point0.X + 0.25 * scale1, y + j * Vincr * Vexag / 5, 0));

                                linie_small_left.Layer = layer_grid_small_lines;
                                linie_small_left.Linetype = "ByLayer";
                                BTrecord.AppendEntity(linie_small_left);
                                Trans1.AddNewlyCreatedDBObject(linie_small_left, true);

                                Line linie_small_right = new Line(new Point3d(EndX - 0.25 * scale1, y + j * Vincr * Vexag / 5, 0),
                                                                  new Point3d(EndX, y + j * Vincr * Vexag / 5, 0));

                                linie_small_right.Layer = layer_grid_small_lines;
                                linie_small_right.Linetype = "ByLayer";
                                BTrecord.AppendEntity(linie_small_right);
                                Trans1.AddNewlyCreatedDBObject(linie_small_right, true);
                            }
                        }

                        Line linie_small_left0 = new Line(new Point3d(Point0.X, y, 0), new Point3d(Point0.X + 0.5 * scale1, y, 0));

                        linie_small_left0.Layer = layer_grid_small_lines;
                        linie_small_left0.Linetype = "ByLayer";
                        BTrecord.AppendEntity(linie_small_left0);
                        Trans1.AddNewlyCreatedDBObject(linie_small_left0, true);

                        Line linie_small_right0 = new Line(new Point3d(EndX - 0.5 * scale1, y, 0), new Point3d(EndX, y, 0));

                        linie_small_right0.Layer = layer_grid_small_lines;
                        linie_small_right0.Linetype = "ByLayer";
                        BTrecord.AppendEntity(linie_small_right0);
                        Trans1.AddNewlyCreatedDBObject(linie_small_right0, true);


                        if (leftElev == true)
                        {
                            MText Mt_el_left = new MText();
                            Mt_el_left.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_left.Layer = Layer_text;
                            Mt_el_left.Attachment = AttachmentPoint.MiddleLeft;
                            Mt_el_left.TextHeight = Texth;
                            Mt_el_left.TextStyleId = Textstyleid;
                            Mt_el_left.Location = new Point3d(Point0.X + 0.25 * scale1, Point0.Y + i * Vincr * Vexag + 0.75 * Texth, 0);
                            BTrecord.AppendEntity(Mt_el_left);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_left, true);


                        }

                        if (rightElev == true)
                        {
                            MText Mt_el_right = new MText();
                            Mt_el_right.Contents = (Downelev + i * Vincr).ToString() + Elev_suffix;
                            Mt_el_right.Layer = Layer_text;
                            Mt_el_right.Attachment = AttachmentPoint.MiddleRight;
                            Mt_el_right.TextHeight = Texth;
                            Mt_el_right.TextStyleId = Textstyleid;
                            Mt_el_right.Location = new Point3d(EndX - 0.25 * scale1, Point0.Y + i * Vincr * Vexag + 0.75 * Texth, 0);
                            BTrecord.AppendEntity(Mt_el_right);
                            Trans1.AddNewlyCreatedDBObject(Mt_el_right, true);



                        }
                    }

                    #endregion


                    #region poly graph
                    Polyline Poly_graph = new Polyline();
                    int idx_p = 0;


                    if (dt_ground != null && dt_ground.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_ground.Rows.Count; ++i)
                        {
                            if (dt_ground.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_ground.Rows[i][Col_elev]);
                                if (dt_ground.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_ground.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - start_grid) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }

                        Poly_graph.Layer = Layer_ground;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);
                    }



                    Poly_graph = new Polyline();
                    idx_p = 0;

                    if (dt_hydrant != null && dt_hydrant.Rows.Count > 1)
                    {
                        for (int i = 0; i < dt_hydrant.Rows.Count; ++i)
                        {
                            if (dt_hydrant.Rows[i][Col_elev] != DBNull.Value)
                            {
                                double z1 = Convert.ToDouble(dt_hydrant.Rows[i][Col_elev]);
                                if (dt_hydrant.Rows[i][Col_sta] != DBNull.Value)
                                {
                                    double Sta1 = Convert.ToDouble(dt_hydrant.Rows[i][Col_sta]);
                                    Point2d ptp = new Point2d(Point0.X + (Sta1 - start_grid) * Hexag, Point0.Y + (z1 - Downelev) * Vexag);
                                    Poly_graph.AddVertexAt(idx_p, ptp, 0, 0, 0);
                                    idx_p = idx_p + 1;
                                }
                            }
                        }

                        Poly_graph.Layer = layer_hydrant;
                        BTrecord.AppendEntity(Poly_graph);
                        Trans1.AddNewlyCreatedDBObject(Poly_graph, true);

                        // MLeader ml_fireh = creaza_mleader(Poly_graph.EndPoint,"test",Texth,6,30);


                    }



                    #endregion



                }

                Trans1.Commit();
            }


        }


        private MLeader creaza_mleader(Point3d pt_ins, string continut, double texth, double delta_x, double delta_y, double lgap, double dogl, double arrow)
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

                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);
                Trans1.Commit();
            }




            return mleader1;







        }



        private void button_load_profile_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_prof_Vspacing.Text) == false)
            {
                MessageBox.Show("please specify the vertical spacing");
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
                        Polyline poly_hydrant = Build_2d_poly_from_dt(dt_pts);



                        string Col_elev = "elev";

                        double vincr = Convert.ToDouble(textBox_prof_Vspacing.Text);


                        dt_prof_hydrant = creaza_dt_sta_and_elev_for_hydrant(poly_hydrant);


                        double Downelev = 0;
                        double Upelev = 0;

                        if (dt_prof_hydrant != null)
                        {
                            if (dt_prof_hydrant.Rows.Count > 2)
                            {
                                double Min_el = 100000;
                                double Max_el = -100000;
                                for (int i = 0; i < dt_prof_hydrant.Rows.Count; ++i)
                                {
                                    if (dt_prof_hydrant.Rows[i][Col_elev] != DBNull.Value)
                                    {
                                        double z1 = Convert.ToDouble(dt_prof_hydrant.Rows[i][Col_elev]);
                                        if (z1 > Max_el) Max_el = z1;
                                        if (z1 < Min_el) Min_el = z1;
                                    }
                                }
                                Downelev = Functions.Round_Down_as_double(Min_el, vincr) - 1 * vincr;
                                Upelev = Functions.Round_Up_as_double(Max_el, vincr) + 2 * vincr;
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



        private void button_load_contours_Click(object sender, EventArgs e)
        {
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
                            dt_prof_cont_ground = null;
                            label_load_contours.Visible = false;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Editor1.SetImpliedSelection(Empty_array);

                            return;
                        }




                        Polyline poly_hydrant = Build_2d_poly_from_dt(dt_pts);



                        if (poly_hydrant != null)
                        {

                            double hincr = Convert.ToDouble(textBox_prof_Hspacing.Text);

                            double Len1 = poly_hydrant.Length;

                            double bear1 = Functions.GET_Bearing_rad(poly_hydrant.GetPoint2dAt(1).X, poly_hydrant.GetPoint2dAt(1).Y, poly_hydrant.GetPoint2dAt(0).X, poly_hydrant.GetPoint2dAt(0).Y);
                            double bear2 = Functions.GET_Bearing_rad(poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 2).X,
                                                                                         poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 2).Y,
                                                                                             poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).X,
                                                                                                 poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).Y);

                            if (Len1 < hincr)
                            {

                                Polyline poly1 = new Polyline();
                                poly1.AddVertexAt(0, poly_hydrant.GetPoint2dAt(0), 0, 0, 0);
                                poly1.AddVertexAt(1, new Point2d(poly_hydrant.GetPoint2dAt(0).X + hincr - 1, poly_hydrant.GetPoint2dAt(0).Y), 0, 0, 0);
                                poly1.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly1.StartPoint));
                                poly_hydrant.AddVertexAt(0, poly1.GetPoint2dAt(1), 0, 0, 0);




                                Polyline poly2 = new Polyline();
                                poly2.AddVertexAt(0, poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1), 0, 0, 0);
                                poly2.AddVertexAt(1, new Point2d(poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).X + hincr - 1 - Len1, poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).Y), 0, 0, 0);
                                poly2.TransformBy(Matrix3d.Rotation(bear2, Vector3d.ZAxis, poly2.StartPoint));
                                poly_hydrant.AddVertexAt(poly_hydrant.NumberOfVertices, poly2.GetPoint2dAt(1), 0, 0, 0);
                            }


                            #region poly contours

                            dt_prof_cont_ground = new System.Data.DataTable();
                            dt_prof_cont_ground.Columns.Add("ptno", typeof(int));
                            dt_prof_cont_ground.Columns.Add("x", typeof(double));
                            dt_prof_cont_ground.Columns.Add("y", typeof(double));
                            dt_prof_cont_ground.Columns.Add("sta", typeof(double));
                            dt_prof_cont_ground.Columns.Add("elev", typeof(double));






                            Polyline poly_start = new Polyline();
                            poly_start.AddVertexAt(0, poly_hydrant.GetPoint2dAt(0), 0, 0, 0);
                            poly_start.AddVertexAt(1, new Point2d(poly_hydrant.GetPoint2dAt(0).X + 1000, poly_hydrant.GetPoint2dAt(0).Y), 0, 0, 0);


                            poly_start.TransformBy(Matrix3d.Rotation(bear1, Vector3d.ZAxis, poly_start.StartPoint));

                            Polyline poly_end = new Polyline();
                            poly_end.AddVertexAt(0, poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1), 0, 0, 0);
                            poly_end.AddVertexAt(1, new Point2d(poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).X + 1000, poly_hydrant.GetPoint2dAt(poly_hydrant.NumberOfVertices - 1).Y), 0, 0, 0);



                            poly_end.TransformBy(Matrix3d.Rotation(bear2, Vector3d.ZAxis, poly_end.StartPoint));






                            double start_elev = -1234.234;
                            double calc_sta_start = 1234.234;
                            double end_elev = -1234.234;
                            double calc_sta_end = 1234.234;



                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline poly_cont = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                Line line_cont = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Line;
                                if (poly_cont != null || line_cont != null)
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
                                        if (poly_cont != null)
                                        {
                                            elev1 = poly_cont.Elevation;
                                        }
                                        if (line_cont != null)
                                        {
                                            elev1 = line_cont.StartPoint.Z;
                                        }

                                    }

                                    Polyline pcnt = new Polyline();
                                    if (poly_cont != null)
                                    {
                                        pcnt = poly_cont.Clone() as Polyline;
                                    }
                                    if (line_cont != null)
                                    {
                                        pcnt.AddVertexAt(0, new Point2d(line_cont.StartPoint.X, line_cont.StartPoint.Y), 0, 0, 0);
                                        pcnt.AddVertexAt(1, new Point2d(line_cont.EndPoint.X, line_cont.EndPoint.Y), 0, 0, 0);
                                    }

                                    pcnt.Elevation = poly_hydrant.Elevation;

                                    Polyline pstart = new Polyline();
                                    pstart = poly_start.Clone() as Polyline;
                                    pstart.Elevation = pcnt.Elevation;

                                    Polyline pend = new Polyline();
                                    pend = poly_end.Clone() as Polyline;
                                    pend.Elevation = pcnt.Elevation;





                                    Point3dCollection colint = Functions.Intersect_on_both_operands(pcnt, poly_hydrant);
                                    if (colint.Count > 0)
                                    {
                                        for (int j = 0; j < colint.Count; ++j)
                                        {
                                            double sta1 = poly_hydrant.GetDistAtPoint(poly_hydrant.GetClosestPointTo(colint[j], Vector3d.ZAxis, false));



                                            dt_prof_cont_ground.Rows.Add();
                                            dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["x"] = colint[j].X;
                                            dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["y"] = colint[j].Y;
                                            dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["elev"] = elev1;
                                            dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["ptno"] = Convert.ToInt32(sta1);
                                            dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["sta"] = sta1;
                                        }
                                    }

                                    Point3dCollection colint_start = new Point3dCollection();
                                    pstart.IntersectWith(pcnt, Intersect.ExtendBoth, colint_start, IntPtr.Zero, IntPtr.Zero);

                                    //Functions.Intersect_on_both_operands(pstart, pcnt);
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


                                    Point3dCollection colint_end = new Point3dCollection();
                                    pend.IntersectWith(pcnt, Intersect.ExtendBoth, colint_end, IntPtr.Zero, IntPtr.Zero);
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
                                }
                            }


                            if (dt_prof_cont_ground.Rows.Count > 0)
                            {
                                dt_prof_cont_ground = Functions.Sort_data_table(dt_prof_cont_ground, "sta");
                            }


                            if (calc_sta_start < 1234.234)
                            {
                                double elev1 = Convert.ToDouble(dt_prof_cont_ground.Rows[0]["elev"]);
                                double sta1 = Convert.ToDouble(dt_prof_cont_ground.Rows[0]["sta"]);


                                double delta_calc = sta1 * Math.Abs(start_elev - elev1) / (sta1 + calc_sta_start);
                                double elev2 = elev1 + delta_calc;
                                if (elev1 > start_elev) elev2 = elev1 - delta_calc;

                                System.Data.DataRow row0 = dt_prof_cont_ground.NewRow();

                                row0["x"] = poly_hydrant.StartPoint.X;
                                row0["y"] = poly_hydrant.StartPoint.Y;
                                row0["elev"] = elev2;
                                row0["ptno"] = 0;
                                row0["sta"] = 0;
                                if (dt_prof_cont_ground.Rows.Count > 0)
                                {
                                    dt_prof_cont_ground.Rows.InsertAt(row0, 0);
                                }
                                else
                                {
                                    dt_prof_cont_ground.ImportRow(row0);
                                }
                            }
                            if (calc_sta_end < 1234.234)
                            {
                                double elev1 = Convert.ToDouble(dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["elev"]);
                                double sta1 = Convert.ToDouble(dt_prof_cont_ground.Rows[dt_prof_cont_ground.Rows.Count - 1]["sta"]);
                                double stax = poly_hydrant.Length;

                                double delta_calc = (Math.Abs(sta1 - stax) * Math.Abs(end_elev - elev1)) / (Math.Abs(sta1 - stax) + calc_sta_end);
                                double elev2 = elev1 + delta_calc;
                                if (elev1 > end_elev) elev2 = elev1 - delta_calc;

                                System.Data.DataRow row0 = dt_prof_cont_ground.NewRow();
                                row0["x"] = poly_hydrant.EndPoint.X;
                                row0["y"] = poly_hydrant.EndPoint.Y;
                                row0["elev"] = elev2;
                                row0["ptno"] = Convert.ToInt32(poly_hydrant.Length);
                                row0["sta"] = stax;

                                if (dt_prof_cont_ground.Rows.Count > 0)
                                {
                                    dt_prof_cont_ground.Rows.InsertAt(row0, dt_prof_cont_ground.Rows.Count);
                                }
                                else
                                {
                                    dt_prof_cont_ground.ImportRow(row0);
                                }

                            }


                            if (Len1 < hincr)
                            {

                                for (int i = 0; i < dt_prof_cont_ground.Rows.Count; ++i)
                                {
                                    dt_prof_cont_ground.Rows[i]["sta"] = hincr - 1 - Convert.ToDouble(dt_prof_cont_ground.Rows[i]["sta"]);
                                }

                            }

                            if (hydrant_is_pozitiv == false && dt_prof_cont_ground.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt_prof_cont_ground.Rows.Count; ++i)
                                {
                                    dt_prof_cont_ground.Rows[i]["sta"] = -Convert.ToDouble(dt_prof_cont_ground.Rows[i]["sta"]);
                                }
                                dt_prof_cont_ground = Functions.Sort_data_table(dt_prof_cont_ground, "sta");
                            }



                        }
                        #endregion
                        //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_prof_cont_tob_ne);
                        Trans1.Commit();
                        label_load_contours.Visible = true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                label_load_contours.Visible = false;
                dt_prof_cont_ground = null;
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }

    }
}
