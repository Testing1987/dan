using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class DivBerm : Form
    {
        private bool clickdragdown;
        private Point lastLocation;
        Polyline Poly_avg1 = null;
        List<ObjectId> List_txt = null;
        System.Data.DataTable dt_slope_ranges = null;
        List<ObjectId> List_poly = null;
        string nume_layer;

        System.Data.DataTable dt_terrain = null;
        System.Data.DataTable dt_blocks = null;
        System.Data.DataTable dt_prof = null;

        Polyline Prof_poly = null;
        Polyline Slope_poly = null;


        string Col_sta = "Station";
        string Col_MMid = "MMID";
        string Col_sta_eq = "StationEq";
        string Col_Type = "Type";
        string Col_Elev = "Elev";
        string Col_Elev1 = "Elev1";
        string Col_Elev2 = "Elev2";

        ObjectId slope_id = ObjectId.Null;
        ObjectId profile_id = ObjectId.Null;

        public DivBerm()
        {
            InitializeComponent();
            nume_layer = "no_plot - [" + Environment.UserName.ToUpper() + "]";
            if (Functions.is_dan_popescu() == true) panel_dan.Visible = true;

        }

        #region minimize and move close

        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
            Poly_avg1 = null;
            List_txt = null;
            dt_slope_ranges = null;
            List_poly = null;

            dt_terrain = null;
            dt_blocks = null;
            dt_prof = null;

            Prof_poly = null;
            Slope_poly = null;
            this.Close();
        }
        private void clickmove_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown = true;
            lastLocation = e.Location;
        }

        private void clickmove_MouseMove(object sender, MouseEventArgs e)
        {
            if (clickdragdown)
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

        #endregion

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_add_div_berm);
            lista_butoane.Add(button_add_vertex);
            lista_butoane.Add(button_average_slope);
            lista_butoane.Add(button_Exit);
            lista_butoane.Add(button_export_data_to_excel);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(button_move_vertex);
            lista_butoane.Add(button_remove_mult_vertices);
            lista_butoane.Add(button_remove_vertex);
            lista_butoane.Add(button_insert_terrain);
            lista_butoane.Add(button_place_div_berms_manualy);
            lista_butoane.Add(button_recalculate_position);
            lista_butoane.Add(button_place_berm_based_on_point);
            lista_butoane.Add(button_place_berm_based_on_sta);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_add_div_berm);
            lista_butoane.Add(button_add_vertex);
            lista_butoane.Add(button_average_slope);
            lista_butoane.Add(button_Exit);
            lista_butoane.Add(button_export_data_to_excel);
            lista_butoane.Add(button_minimize);
            lista_butoane.Add(button_move_vertex);
            lista_butoane.Add(button_remove_mult_vertices);
            lista_butoane.Add(button_remove_vertex);
            lista_butoane.Add(button_insert_terrain);
            lista_butoane.Add(button_place_div_berms_manualy);
            lista_butoane.Add(button_recalculate_position);
            lista_butoane.Add(button_place_berm_based_on_point);
            lista_butoane.Add(button_place_berm_based_on_sta);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private System.Data.DataTable Create_slope_ranges()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("start", typeof(double));
            dt1.Columns.Add("end", typeof(double));
            dt1.Columns.Add("label", typeof(string));

            string ss1 = textBox_ss1.Text;
            string se1 = textBox_se1.Text;

            dt1.Rows.Add();
            if (Functions.IsNumeric(ss1) == true && Functions.IsNumeric(se1) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss1));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se1));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss1 + "-" + se1 + "%";
            }



            string ss2 = textBox_ss2.Text;
            string se2 = textBox_se2.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss2) == true && Functions.IsNumeric(se2) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss2));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se2));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss2 + "-" + se2 + "%";
            }



            string ss3 = textBox_ss3.Text;
            string se3 = textBox_se3.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss3) == true && Functions.IsNumeric(se3) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss3));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se3));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss3 + "-" + se3 + "%";
            }



            string ss4 = textBox_ss4.Text;
            string se4 = textBox_se4.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss4) == true && Functions.IsNumeric(se4) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss4));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se4));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss4 + "-" + se4 + "%";
            }



            string ss5 = textBox_ss5.Text;
            string se5 = textBox_se5.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss5) == true && Functions.IsNumeric(se5) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss5));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se5));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss5 + "-" + se5 + "%";
            }



            string ss6 = textBox_ss6.Text;
            string se6 = textBox_se6.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss6) == true && Functions.IsNumeric(se6) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss6));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se6));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss6 + "-" + se6 + "%";
            }



            string ss7 = textBox_ss7.Text;
            string se7 = textBox_se7.Text;
            dt1.Rows.Add();
            if (Functions.IsNumeric(ss7) == true && Functions.IsNumeric(se7) == true)
            {
                dt1.Rows[dt1.Rows.Count - 1]["start"] = Math.Abs(Convert.ToDouble(ss7));
                dt1.Rows[dt1.Rows.Count - 1]["end"] = Math.Abs(Convert.ToDouble(se7));
                dt1.Rows[dt1.Rows.Count - 1]["label"] = ss7 + "-" + se7 + "%";
            }



            return dt1;
        }

        private double calc_slope(Point3d pt1, Point3d pt2)
        {
            double DeltaX = pt2.X - pt1.X;
            double DeltaY = pt2.Y - pt1.Y;
            return Math.Round(100 * DeltaY / DeltaX, 1);
        }

        private bool is_the_same_sign(double nr1, double nr2)
        {
            if (nr1 != 0 && nr2 != 0)
            {
                double no1 = nr1 / Math.Abs(nr1);
                double no2 = nr2 / Math.Abs(nr2);
                if (no1 == no2)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            if (nr1 == 0 && nr2 == 0)
            {
                return true;
            }

            return false;
        }

        private System.Data.DataTable creaza_dt_from_poly(Polyline Poly1)
        {

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("slope", typeof(double));
            dt1.Columns.Add("dist_ahead", typeof(double));
            dt1.Columns.Add("rotation", typeof(double));
            dt1.Columns.Add("texth", typeof(double));
            dt1.Columns.Add("index", typeof(int));
            dt1.Columns.Add("add_vertex", typeof(int));
            dt1.Columns.Add("x", typeof(double));
            dt1.Columns.Add("y", typeof(double));

            for (int i = 0; i < Poly1.NumberOfVertices - 1; ++i)
            {
                double x1 = Poly1.GetPointAtParameter(i).X;
                double y1 = Poly1.GetPointAtParameter(i).Y;
                double x2 = Poly1.GetPointAtParameter(i + 1).X;
                double y2 = Poly1.GetPointAtParameter(i + 1).Y;
                double DeltaX = x2 - x1;
                double DeltaY = y2 - y1;
                double Slope = Math.Round(100 * DeltaY / DeltaX, 1);
                double Rot1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                double Dist = Math.Abs(x1 - x2);

                #region calculate text height

                double Texth = 3;
                string plus = "";
                if (Slope > 0) plus = "+";
                double Textwidth;

                do
                {
                    Texth = Texth - 0.1;
                    if (Texth <= 0)
                    {
                        Texth = 0.5;
                        Dist = 100;
                    }
                    MText MTT = new MText();
                    MTT.Contents = plus + Functions.Get_String_Rounded(Slope, 1) + "%";
                    MTT.Attachment = AttachmentPoint.BottomCenter;
                    MTT.Location = new Point3d(0, 0, 0);
                    MTT.TextHeight = Texth;
                    MTT.Rotation = 0;

                    Extents3d Extend1 = MTT.GeometricExtents;
                    Textwidth = Math.Abs(Extend1.MaxPoint.X - Extend1.MinPoint.X);

                } while (Textwidth + 2 > Dist);

                #endregion

                dt1.Rows.Add();
                dt1.Rows[dt1.Rows.Count - 1]["slope"] = Slope;
                dt1.Rows[dt1.Rows.Count - 1]["dist_ahead"] = Dist;
                dt1.Rows[dt1.Rows.Count - 1]["rotation"] = Rot1;
                dt1.Rows[dt1.Rows.Count - 1]["texth"] = Texth;
                dt1.Rows[dt1.Rows.Count - 1]["index"] = i;
                dt1.Rows[dt1.Rows.Count - 1]["x"] = x1;
                dt1.Rows[dt1.Rows.Count - 1]["y"] = y1;
                if (i == Poly1.NumberOfVertices - 2)
                {
                    dt1.Rows.Add();
                    dt1.Rows[dt1.Rows.Count - 1]["slope"] = 400;
                    dt1.Rows[dt1.Rows.Count - 1]["dist_ahead"] = 1;
                    dt1.Rows[dt1.Rows.Count - 1]["rotation"] = 0;
                    dt1.Rows[dt1.Rows.Count - 1]["texth"] = 1;
                    dt1.Rows[dt1.Rows.Count - 1]["index"] = i + 1;
                    dt1.Rows[dt1.Rows.Count - 1]["x"] = x2;
                    dt1.Rows[dt1.Rows.Count - 1]["y"] = y2;
                }
            }
            return dt1;
        }

        private void button_calc_average_slope_Click(object sender, EventArgs e)
        {
            double min_dist = -1;
            string distmin_string = textBox_min_slope.Text;

            if (Functions.IsNumeric(distmin_string) == true)
            {
                min_dist = Convert.ToDouble(distmin_string);
            }

            if (min_dist <= 0)
            {
                MessageBox.Show("minimum distance is not specified properly");
                return;
            }

            if (Functions.IsNumeric(textBox_tolerance.Text) == false)
            {
                MessageBox.Show("tolerance is not specified properly");
                return;
            }



            this.WindowState = FormWindowState.Minimized;

            set_enable_false();
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
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsCL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect profile polyline:");
                        Prompt_optionsCL.SetRejectMessage("\nYou did not selected a lightweight polyline");
                        Prompt_optionsCL.AddAllowedClass(typeof(Polyline), true);
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_CL = Editor1.GetEntity(Prompt_optionsCL);
                        if (Rezultat_CL.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            this.WindowState = FormWindowState.Normal;
                            return;
                        }

                        this.WindowState = FormWindowState.Normal;
                        Prof_poly = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead) as Polyline;

                        if (Prof_poly == null)
                        {
                            MessageBox.Show("you did not select a polyline");
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            this.WindowState = FormWindowState.Normal;
                            return;
                        }

                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        dt_slope_ranges = Create_slope_ranges();

                        List_txt = new List<ObjectId>();
                        Functions.Creaza_layer(nume_layer, 30, false);

                        System.Data.DataTable dt1 = creaza_dt_from_poly(Prof_poly);

                        #region create profile original labels

                        for (int i = 0; i < Prof_poly.NumberOfVertices - 1; ++i)
                        {
                            double x1 = Prof_poly.GetPointAtParameter(i).X;
                            double y1 = Prof_poly.GetPointAtParameter(i).Y;
                            double x2 = Prof_poly.GetPointAtParameter(i + 1).X;
                            double y2 = Prof_poly.GetPointAtParameter(i + 1).Y;

                            string plus = "";
                            double Slope = Convert.ToDouble(dt1.Rows[i]["slope"]);
                            if (Slope > 0) plus = "+";

                            double Texth = Convert.ToDouble(dt1.Rows[i]["texth"]);
                            double Rot1 = Convert.ToDouble(dt1.Rows[i]["rotation"]);

                            MText mt1 = new MText();
                            mt1.Contents = plus + Functions.Get_String_Rounded(Slope, 1) + "%";
                            mt1.Attachment = AttachmentPoint.BottomCenter;
                            mt1.Location = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                            mt1.TextHeight = Texth;
                            mt1.Rotation = Rot1;
                            mt1.Layer = nume_layer;
                            mt1.ColorIndex = 256;
                            BTrecord.AppendEntity(mt1);
                            Trans1.AddNewlyCreatedDBObject(mt1, true);
                        }

                        #endregion



                        Polyline Poly_avg_1_step = new Polyline();
                        Poly_avg_1_step.AddVertexAt(0, new Point2d(Prof_poly.StartPoint.X, Prof_poly.StartPoint.Y), 0, 0, 0);

                        Point3d pt1 = Prof_poly.StartPoint;
                        int j = 1;
                        double offset1 = Convert.ToDouble(textBox_tolerance.Text) / 2;
                        DBObjectCollection dbcol1 = Prof_poly.GetOffsetCurves(offset1);
                        Polyline Poly_down = new Polyline();
                        Poly_down = dbcol1[0] as Polyline;
                        DBObjectCollection dbcol2 = Prof_poly.GetOffsetCurves(-offset1);
                        Polyline Poly_up = new Polyline();
                        Poly_up = dbcol2[0] as Polyline;

                        for (int i = 1; i < dt1.Rows.Count; ++i)
                        {

                            Point3d pt2 = Prof_poly.GetPointAtParameter(i - 1);
                            Point3d pt3 = Prof_poly.GetPointAtParameter(i);

                            Line l3 = new Line(pt1, pt3);
                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(l3, Poly_up);
                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(l3, Poly_down);
                            if (colint1.Count > 0 || colint2.Count > 0)
                            {

                                Poly_avg_1_step.AddVertexAt(j, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                pt1 = pt2;
                                ++j;
                            }

                            if (i == dt1.Rows.Count - 1)
                            {

                                Point3d pt4 = Prof_poly.GetPointAtParameter(Convert.ToInt32(Prof_poly.EndParam));
                                Line l4 = new Line(pt1, pt4);

                                Point3dCollection colint11 = Functions.Intersect_on_both_operands(l4, Poly_up);
                                Point3dCollection colint22 = Functions.Intersect_on_both_operands(l4, Poly_down);
                                if (colint11.Count > 0 || colint22.Count > 0)
                                {

                                    Poly_avg_1_step.AddVertexAt(j, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                    ++j;

                                }
                                Poly_avg_1_step.AddVertexAt(j, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                            }
                        }

                        System.Data.DataTable dt2 = creaza_dt_from_poly(Poly_avg_1_step);

                        double slope1 = Convert.ToDouble(dt2.Rows[0]["slope"]);
                        dt2.Rows[0]["add_vertex"] = 1;

                        for (int i = 1; i < dt2.Rows.Count - 1; ++i)
                        {
                            double slope2 = Convert.ToDouble(dt2.Rows[i]["slope"]);
                            double cat1 = find_category(dt_slope_ranges, slope1);
                            double cat2 = find_category(dt_slope_ranges, slope2);
                            if (cat1 == cat2)
                            {
                                if (slope1 == 0 && cat2 == 1)
                                {
                                    dt2.Rows[i]["add_vertex"] = 0;
                                }
                                else if (slope2 == 0 && cat1 == 1)
                                {
                                    dt2.Rows[i]["add_vertex"] = 0;
                                }
                                else if (is_the_same_sign(slope1, slope2) == true)
                                {
                                    dt2.Rows[i]["add_vertex"] = 0;
                                }
                                else
                                {
                                    dt2.Rows[i]["add_vertex"] = 1;
                                }

                            }
                            else
                            {
                                dt2.Rows[i]["add_vertex"] = 1;
                            }

                            slope1 = slope2;
                        }


                        dt2.Rows[dt2.Rows.Count - 1]["add_vertex"] = 1;

                        for (int i = dt2.Rows.Count - 2; i > 0; --i)
                        {
                            int xadd = Convert.ToInt32(dt2.Rows[i]["add_vertex"]);
                            if (xadd == 0)
                            {
                                dt2.Rows[i].Delete();
                            }
                        }

                        if (checkBox_use_min_dist.Checked == true)
                        {
                            for (int i = 1; i < dt2.Rows.Count; ++i)
                            {
                                double x1 = Convert.ToDouble(dt2.Rows[i - 1]["x"]);
                                double y1 = Convert.ToDouble(dt2.Rows[i - 1]["y"]);
                                double x2 = Convert.ToDouble(dt2.Rows[i]["x"]);
                                double y2 = Convert.ToDouble(dt2.Rows[i]["y"]);
                                double dist1 = Math.Abs(x1 - x2);
                                if (dist1 < min_dist)
                                {
                                    dt2.Rows[i]["add_vertex"] = 0;
                                    for (int k = i + 1; k < dt2.Rows.Count - 1; ++k)
                                    {
                                        x2 = Convert.ToDouble(dt2.Rows[k]["x"]);
                                        y2 = Convert.ToDouble(dt2.Rows[k]["y"]);
                                        dist1 = Math.Abs(x1 - x2);
                                        if (dist1 < min_dist)
                                        {
                                            dt2.Rows[k]["add_vertex"] = 0;
                                        }
                                        else
                                        {
                                            i = k - 1;
                                            k = dt2.Rows.Count;
                                        }
                                    }

                                }

                            }
                        }

                        List_poly = new List<ObjectId>();

                        Polyline Poly_avg_2_step = new Polyline();

                        j = 0;

                        for (int i = 0; i < dt2.Rows.Count; ++i)
                        {
                            double x = Convert.ToDouble(dt2.Rows[i]["x"]);
                            double y = Convert.ToDouble(dt2.Rows[i]["y"]);
                            int xadd = Convert.ToInt32(dt2.Rows[i]["add_vertex"]);
                            if (xadd == 1)
                            {

                                Poly_avg_2_step.AddVertexAt(j, new Point2d(x, y), 0, 0, 0);
                                if (i > 0)
                                {
                                    for (int m = i - 1; m >= 0; --m)
                                    {
                                        int xadd0 = Convert.ToInt32(dt2.Rows[m]["add_vertex"]);
                                        if (xadd0 == 1)
                                        {
                                            double x0 = Convert.ToDouble(dt2.Rows[m]["x"]);
                                            double y0 = Convert.ToDouble(dt2.Rows[m]["y"]);
                                            double slope2 = calc_slope(new Point3d(x0, y0, 0), new Point3d(x, y, 0));
                                            int cat2 = find_category(dt_slope_ranges, slope2);

                                            if (cat2 != -1)
                                            {
                                                double dist1 = Math.Abs(x0 - x);
                                                double Rot1 = Functions.GET_Bearing_rad(x0, y0, x, y);
                                                string plus = "";
                                                if (slope2 > 0) plus = "+";

                                                MText mt1 = new MText();
                                                mt1.Contents = plus + Functions.Get_String_Rounded(slope2, 1) + "% - Rge " + cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(dist1, 1);
                                                mt1.Attachment = AttachmentPoint.TopCenter;
                                                mt1.Location = new Point3d((x0 + x) / 2, (y0 + y) / 2, 0);
                                                mt1.TextHeight = 1;
                                                mt1.Rotation = Rot1;
                                                mt1.Layer = nume_layer;
                                                mt1.ColorIndex = color_index(cat2);
                                                BTrecord.AppendEntity(mt1);
                                                Trans1.AddNewlyCreatedDBObject(mt1, true);

                                                List_txt.Add(mt1.ObjectId);
                                            }
                                            else
                                            {
                                                List_txt.Add(ObjectId.Null);
                                            }

                                            m = -1;
                                        }
                                    }
                                }
                                ++j;
                            }
                        }


                        Poly_avg_2_step.ColorIndex = 7;
                        Poly_avg_2_step.Layer = nume_layer;
                        BTrecord.AppendEntity(Poly_avg_2_step);
                        Trans1.AddNewlyCreatedDBObject(Poly_avg_2_step, true);




                        Poly_avg1 = Poly_avg_2_step;

                        if (List_poly.Count > 0)
                        {
                            DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                            ObjectIdCollection col1 = new ObjectIdCollection();
                            for (int i = 0; i < List_poly.Count; ++i)
                            {
                                col1.Add(List_poly[i]);
                            }

                            DrawOrderTable1.MoveToBottom(col1);
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
            this.WindowState = FormWindowState.Normal;
        }
        private int find_category(System.Data.DataTable dt2, double slope1)
        {
            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                if (dt2.Rows[j]["start"] != DBNull.Value && dt2.Rows[j]["end"] != DBNull.Value)
                {
                    double ss1 = Convert.ToDouble(dt2.Rows[j]["start"]);
                    double se1 = Convert.ToDouble(dt2.Rows[j]["end"]);

                    if (Math.Abs(slope1) >= ss1 && Math.Abs(slope1) < se1)
                    {
                        return j + 1;
                    }
                }
            }
            return -1;
        }


        private int color_index(int cat1)
        {
            int ci = 256;
            switch (cat1)
            {
                case 1:
                    ci = 2;
                    break;
                case 2:
                    ci = 3;
                    break;
                case 3:
                    ci = 6;
                    break;
                case 4:
                    ci = 1;
                    break;
                case 5:
                    ci = 5;
                    break;
                case 6:
                    ci = 4;
                    break;
                case 7:
                    ci = 7;
                    break;
                default:
                    ci = 9;
                    break;
            }
            return ci;
        }



        private void add_ids_to_list_txt(Transaction Trans1, BlockTableRecord BTrecord, Polyline Poly2)
        {
            foreach (ObjectId id1 in BTrecord)
            {
                MText sMtext = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
                if (sMtext != null)
                {
                    if (sMtext.Layer.ToLower().Contains("no_plot") == true)
                    {
                        Point3d point_mt = sMtext.Location;
                        Point3d pt_on_poly = Poly2.GetClosestPointTo(point_mt, Vector3d.ZAxis, false);
                        if (point_mt.DistanceTo(point_mt) <= 0.1)
                        {
                            double param1 = Poly2.GetParameterAtPoint(pt_on_poly);
                            int k = Convert.ToInt32(Math.Floor(param1));
                            if (Math.Round(param1 - k, 1) == 0.5)
                            {
                                List_txt[k] = sMtext.ObjectId;
                            }

                        }


                    }
                }
            }
        }


        private void button_add_vertex_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        this.WindowState = FormWindowState.Minimized;

                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);


                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            Polyline Poly2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly_avg1 != null)
                            {
                                if (Poly2.ObjectId != Poly_avg1.ObjectId)
                                {

                                    List_txt = new List<ObjectId>();


                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);

                                }
                                else
                                {
                                    if (List_txt == null)
                                    {
                                        List_txt = new List<ObjectId>();
                                    }

                                    if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                    {
                                        List_txt.Clear();

                                        for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                        {
                                            List_txt.Add(ObjectId.Null);
                                        }
                                        add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                    }
                                }

                            }
                            else
                            {
                                if (List_txt == null)
                                {
                                    List_txt = new List<ObjectId>();
                                }

                                if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                {
                                    List_txt.Clear();

                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                }
                            }

                            Poly_avg1 = Poly2;

                        }



                        if (Poly_avg1 != null)
                        {
                            dt_slope_ranges = Create_slope_ranges();
                            Polyline Poly1 = Trans1.GetObject(Poly_avg1.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly1 != null)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d Point1 = Point_res1.Value;
                                Point1 = new Point3d(Point1.X, Point1.Y, Poly1.Elevation);
                                Point3d point_on_poly = Poly1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                                double param1 = Poly1.GetParameterAtPoint(point_on_poly);

                                int index1 = Convert.ToInt32(Math.Floor(param1));
                                int index2 = Convert.ToInt32(Math.Ceiling(param1));
                                if (index2 < Poly1.EndParam && index2 > index1)
                                {
                                    Poly1.AddVertexAt(index2, new Point2d(Point1.X, Point1.Y), 0, 0, 0);

                                    if (List_txt[index1] != ObjectId.Null)
                                    {
                                        try
                                        {
                                            MText existing_label = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                            existing_label.Erase();
                                        }
                                        catch (System.Exception)
                                        {

                                        }

                                    }
                                    List_txt.RemoveAt(index1);

                                    LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1);

                                    Point2d pt1 = segm1.StartPoint;
                                    Point2d pt2 = segm1.EndPoint;

                                    double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    double Texth = 1;
                                    double Dist = Math.Abs(pt2.X - pt1.X);
                                    double x12 = pt1.X;
                                    double y12 = pt1.Y;
                                    double x22 = pt2.X;
                                    double y22 = pt2.Y;
                                    double DeltaX2 = x22 - x12;
                                    double DeltaY2 = y22 - y12;
                                    double Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                    int Cat2 = find_category(dt_slope_ranges, Slope2);

                                    string plus = "";
                                    if (Slope2 > 0) plus = "+";

                                    if (Cat2 != -1)
                                    {
                                        MText mt1 = new MText();
                                        mt1.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                        mt1.Attachment = AttachmentPoint.TopCenter;
                                        mt1.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                        mt1.TextHeight = Texth;
                                        mt1.Rotation = Rot1;
                                        mt1.Layer = nume_layer;
                                        mt1.ColorIndex = color_index(Cat2);
                                        BTrecord.AppendEntity(mt1);
                                        Trans1.AddNewlyCreatedDBObject(mt1, true);
                                        List_txt.Insert(index1, mt1.ObjectId);
                                    }
                                    else
                                    {
                                        List_txt.Insert(index1, ObjectId.Null);
                                    }




                                    segm1 = Poly1.GetLineSegment2dAt(index2);

                                    pt1 = segm1.StartPoint;
                                    pt2 = segm1.EndPoint;

                                    Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    Dist = Math.Abs(pt1.X - pt2.X);
                                    x12 = pt1.X;
                                    y12 = pt1.Y;
                                    x22 = pt2.X;
                                    y22 = pt2.Y;
                                    DeltaX2 = x22 - x12;
                                    DeltaY2 = y22 - y12;
                                    Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                    Cat2 = find_category(dt_slope_ranges, Slope2);
                                    plus = "";
                                    if (Slope2 > 0) plus = "+";


                                    if (Cat2 != -1)
                                    {
                                        MText mt2 = new MText();
                                        mt2.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                        mt2.Attachment = AttachmentPoint.TopCenter;
                                        mt2.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                        mt2.TextHeight = Texth;
                                        mt2.Rotation = Rot1;
                                        mt2.Layer = nume_layer;
                                        mt2.ColorIndex = color_index(Cat2);
                                        BTrecord.AppendEntity(mt2);
                                        Trans1.AddNewlyCreatedDBObject(mt2, true);
                                        List_txt.Insert(index2, mt2.ObjectId);
                                    }
                                    else
                                    {
                                        List_txt.Insert(index2, ObjectId.Null);
                                    }
                                }

                                Update_blocks_slopeID(Trans1, BTrecord, Poly_avg1);

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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }



        private void button_remove_vertex_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        this.WindowState = FormWindowState.Minimized;

                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            Polyline Poly2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly_avg1 != null)
                            {
                                if (Poly2.ObjectId != Poly_avg1.ObjectId)
                                {
                                    List_txt = new List<ObjectId>();
                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                }
                                else
                                {
                                    if (List_txt == null)
                                    {
                                        List_txt = new List<ObjectId>();
                                    }

                                    if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                    {
                                        List_txt.Clear();

                                        for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                        {
                                            List_txt.Add(ObjectId.Null);
                                        }
                                        add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                    }
                                }

                            }
                            else
                            {
                                if (List_txt == null)
                                {
                                    List_txt = new List<ObjectId>();
                                }

                                if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                {
                                    List_txt.Clear();

                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                }
                            }

                            Poly_avg1 = Poly2;

                        }
                        if (Poly_avg1 != null)
                        {
                            Polyline Poly1 = Trans1.GetObject(Poly_avg1.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly1 != null)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify vertex:");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }
                                dt_slope_ranges = Create_slope_ranges();

                                Point3d Point1 = Point_res1.Value;
                                Point1 = new Point3d(Point1.X, Point1.Y, Poly1.Elevation);
                                Point3d point_on_poly = Poly1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                                double param1 = Poly1.GetParameterAtPoint(point_on_poly);

                                int index1 = Convert.ToInt32(Math.Round(param1, 0)) - 1;
                                int index2 = index1 + 1;

                                if (index2 < Poly1.EndParam && index1 > -1)
                                {
                                    Poly1.RemoveVertexAt(index2);

                                    if (List_txt[index2] != ObjectId.Null)
                                    {
                                        try
                                        {
                                            MText existing_label2 = Trans1.GetObject(List_txt[index2], OpenMode.ForWrite) as MText;
                                            existing_label2.Erase();
                                        }
                                        catch (System.Exception)
                                        {

                                        }


                                    }
                                    List_txt.RemoveAt(index2);

                                    if (List_txt[index1] != ObjectId.Null)
                                    {
                                        try
                                        {
                                            MText existing_label1 = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                            existing_label1.Erase();
                                        }
                                        catch (System.Exception)
                                        {

                                        }

                                    }
                                    List_txt.RemoveAt(index1);

                                    LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1);

                                    Point2d pt1 = segm1.StartPoint;
                                    Point2d pt2 = segm1.EndPoint;

                                    double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    double Texth = 1;
                                    double Dist = Math.Abs(pt1.X - pt2.X);
                                    double x12 = pt1.X;
                                    double y12 = pt1.Y;
                                    double x22 = pt2.X;
                                    double y22 = pt2.Y;
                                    double DeltaX2 = x22 - x12;
                                    double DeltaY2 = y22 - y12;
                                    double Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                    int Cat2 = find_category(dt_slope_ranges, Slope2);

                                    string plus = "";
                                    if (Slope2 > 0) plus = "+";

                                    if (Cat2 != -1)
                                    {
                                        MText mt1 = new MText();
                                        mt1.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                        mt1.Attachment = AttachmentPoint.TopCenter;
                                        mt1.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                        mt1.TextHeight = Texth;
                                        mt1.Rotation = Rot1;
                                        mt1.Layer = nume_layer;
                                        mt1.ColorIndex = color_index(Cat2);

                                        BTrecord.AppendEntity(mt1);
                                        Trans1.AddNewlyCreatedDBObject(mt1, true);
                                        List_txt.Insert(index1, mt1.ObjectId);
                                    }
                                    else
                                    {
                                        List_txt.Insert(index1, ObjectId.Null);
                                    }

                                }

                                Update_blocks_slopeID(Trans1, BTrecord, Poly_avg1);



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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();

        }
        private void button_move_vertex_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            dt_slope_ranges = Create_slope_ranges();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                        this.WindowState = FormWindowState.Minimized;

                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);


                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            Polyline Poly2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly_avg1 != null)
                            {
                                if (Poly2.ObjectId != Poly_avg1.ObjectId)
                                {

                                    List_txt = new List<ObjectId>();


                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);

                                }
                                else
                                {
                                    if (List_txt == null)
                                    {
                                        List_txt = new List<ObjectId>();
                                    }

                                    if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                    {
                                        List_txt.Clear();

                                        for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                        {
                                            List_txt.Add(ObjectId.Null);
                                        }
                                        add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                    }
                                }

                            }
                            else
                            {
                                if (List_txt == null)
                                {
                                    List_txt = new List<ObjectId>();
                                }

                                if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                {
                                    List_txt.Clear();

                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                }
                            }

                            Poly_avg1 = Poly2;

                        }
                        if (Poly_avg1 != null)
                        {
                            Polyline Poly1 = Trans1.GetObject(Poly_avg1.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly1 != null)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify vertex:");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d Point1 = Point_res1.Value;
                                Point1 = new Point3d(Point1.X, Point1.Y, Poly1.Elevation);
                                Point3d point_on_poly = Poly1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                                double param1 = Poly1.GetParameterAtPoint(point_on_poly);

                                int index1 = Convert.ToInt32(Math.Round(param1, 0));


                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify new vertex position:");
                                PP2.AllowNone = false;
                                PP2.UseBasePoint = true;
                                PP2.BasePoint = Poly1.GetPointAtParameter(index1);
                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d new_point = Point_res2.Value;

                                Poly1.RemoveVertexAt(index1);

                                Poly1.AddVertexAt(index1, new Point2d(new_point.X, new_point.Y), 0, 0, 0);

                                if (List_txt[index1] != ObjectId.Null)
                                {
                                    try
                                    {
                                        MText existing_label2 = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                        existing_label2.Erase();
                                    }
                                    catch (System.Exception)
                                    {

                                    }

                                }


                                List_txt.RemoveAt(index1);
                                if (List_txt[index1 - 1] != ObjectId.Null)
                                {
                                    try
                                    {
                                        MText existing_label1 = Trans1.GetObject(List_txt[index1 - 1], OpenMode.ForWrite) as MText;
                                        existing_label1.Erase();
                                    }
                                    catch (System.Exception)
                                    {

                                    }
                                }

                                List_txt.RemoveAt(index1 - 1);

                                LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1 - 1);

                                Point2d pt1 = segm1.StartPoint;
                                Point2d pt2 = segm1.EndPoint;

                                double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                double Texth = 1;
                                double Dist = Math.Abs(pt1.X - pt2.X);
                                double x12 = pt1.X;
                                double y12 = pt1.Y;
                                double x22 = pt2.X;
                                double y22 = pt2.Y;
                                double DeltaX2 = x22 - x12;
                                double DeltaY2 = y22 - y12;
                                double Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                int Cat2 = find_category(dt_slope_ranges, Slope2);

                                string plus = "";
                                if (Slope2 > 0) plus = "+";
                                if (Cat2 != -1)
                                {
                                    MText mt1 = new MText();
                                    mt1.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                    mt1.Attachment = AttachmentPoint.TopCenter;
                                    mt1.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                    mt1.TextHeight = Texth;
                                    mt1.Rotation = Rot1;
                                    mt1.Layer = nume_layer;
                                    mt1.ColorIndex = color_index(Cat2);

                                    BTrecord.AppendEntity(mt1);
                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                    List_txt.Insert(index1 - 1, mt1.ObjectId);
                                }
                                else
                                {
                                    List_txt.Insert(index1 - 1, ObjectId.Null);
                                }
                                segm1 = Poly1.GetLineSegment2dAt(index1);

                                pt1 = segm1.StartPoint;
                                pt2 = segm1.EndPoint;

                                Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);

                                Dist = Math.Abs(pt1.X - pt2.X);
                                x12 = pt1.X;
                                y12 = pt1.Y;
                                x22 = pt2.X;
                                y22 = pt2.Y;
                                DeltaX2 = x22 - x12;
                                DeltaY2 = y22 - y12;
                                Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                Cat2 = find_category(dt_slope_ranges, Slope2);

                                plus = "";
                                if (Slope2 > 0) plus = "+";
                                if (Cat2 != -1)
                                {
                                    MText mt2 = new MText();
                                    mt2.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                    mt2.Attachment = AttachmentPoint.TopCenter;
                                    mt2.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                    mt2.TextHeight = Texth;
                                    mt2.Rotation = Rot1;
                                    mt2.Layer = nume_layer;
                                    mt2.ColorIndex = color_index(Cat2);

                                    BTrecord.AppendEntity(mt2);
                                    Trans1.AddNewlyCreatedDBObject(mt2, true);
                                    List_txt.Insert(index1, mt2.ObjectId);
                                }
                                else
                                {
                                    List_txt.Insert(index1, ObjectId.Null);
                                }

                                Update_blocks_slopeID(Trans1, BTrecord, Poly_avg1);


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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }
        private void button_remove_mult_vertices_Click(object sender, EventArgs e)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        this.WindowState = FormWindowState.Minimized;
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);


                        if (Rezultat_centerline.Status == PromptStatus.OK)
                        {
                            Polyline Poly2 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly_avg1 != null)
                            {
                                if (Poly2.ObjectId != Poly_avg1.ObjectId)
                                {
                                    List_txt = new List<ObjectId>();
                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);

                                }
                                else
                                {
                                    if (List_txt == null)
                                    {
                                        List_txt = new List<ObjectId>();
                                    }
                                    if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                    {
                                        List_txt.Clear();

                                        for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                        {
                                            List_txt.Add(ObjectId.Null);
                                        }
                                        add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                    }
                                }
                            }
                            else
                            {
                                if (List_txt == null)
                                {
                                    List_txt = new List<ObjectId>();
                                }

                                if (List_txt.Count != Poly2.NumberOfVertices - 1)
                                {
                                    List_txt.Clear();

                                    for (int i = 0; i < Poly2.NumberOfVertices - 1; ++i)
                                    {
                                        List_txt.Add(ObjectId.Null);
                                    }
                                    add_ids_to_list_txt(Trans1, BTrecord, Poly2);
                                }
                            }
                            Poly_avg1 = Poly2;
                        }
                        if (Poly_avg1 != null)
                        {
                            Polyline Poly1 = Trans1.GetObject(Poly_avg1.ObjectId, OpenMode.ForWrite) as Polyline;
                            if (Poly1 != null)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point1:");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }
                                dt_slope_ranges = Create_slope_ranges();
                                Point3d Point1 = Point_res1.Value;
                                Point1 = new Point3d(Point1.X, Point1.Y, Poly1.Elevation);
                                Point3d point_on_poly1 = Poly1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                                double param1 = Poly1.GetParameterAtPoint(point_on_poly1);

                                int index1 = Convert.ToInt32(Math.Round(param1, 0));

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                                PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point2:");
                                PP2.AllowNone = false;
                                PP2.UseBasePoint = true;
                                PP2.BasePoint = point_on_poly1;
                                Point_res2 = Editor1.GetPoint(PP2);

                                if (Point_res2.Status != PromptStatus.OK)
                                {
                                    Editor1.SetImpliedSelection(Empty_array);
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Point3d Point2 = Point_res2.Value;
                                Point2 = new Point3d(Point2.X, Point2.Y, Poly1.Elevation);
                                Point3d point_on_poly2 = Poly1.GetClosestPointTo(Point2, Vector3d.ZAxis, false);
                                double param2 = Poly1.GetParameterAtPoint(point_on_poly2);

                                int index2 = Convert.ToInt32(Math.Round(param2, 0));

                                if (index1 > index2)
                                {
                                    int t = index1;
                                    index1 = index2;
                                    index2 = t;
                                }

                                for (int i = index2; i >= index1; --i)
                                {
                                    Poly1.RemoveVertexAt(i);

                                    if (index1 == index2)
                                    {
                                        if (List_txt[i - 1] != ObjectId.Null)
                                        {
                                            try
                                            {
                                                MText existing_label2 = Trans1.GetObject(List_txt[i - 1], OpenMode.ForWrite) as MText;
                                                existing_label2.Erase();
                                            }
                                            catch (System.Exception)
                                            {

                                            }

                                        }
                                        List_txt.RemoveAt(i - 1);
                                    }
                                    else
                                    {
                                        if (List_txt[i] != ObjectId.Null)
                                        {
                                            try
                                            {
                                                MText existing_label2 = Trans1.GetObject(List_txt[i], OpenMode.ForWrite) as MText;
                                                existing_label2.Erase();
                                            }
                                            catch (System.Exception)
                                            {

                                            }

                                        }
                                        List_txt.RemoveAt(i);
                                    }
                                }


                                if (List_txt[index1 - 1] != ObjectId.Null)
                                {
                                    try
                                    {
                                        MText existing_label3 = Trans1.GetObject(List_txt[index1 - 1], OpenMode.ForWrite) as MText;
                                        existing_label3.Erase();
                                    }
                                    catch (System.Exception)
                                    {

                                    }


                                }
                                List_txt.RemoveAt(index1 - 1);

                                LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1 - 1);

                                Point2d pt1 = segm1.StartPoint;
                                Point2d pt2 = segm1.EndPoint;

                                double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                double Texth = 1;
                                double Dist = Math.Abs(pt1.X - pt2.X);
                                double x12 = pt1.X;
                                double y12 = pt1.Y;
                                double x22 = pt2.X;
                                double y22 = pt2.Y;
                                double DeltaX2 = x22 - x12;
                                double DeltaY2 = y22 - y12;
                                double Slope2 = Math.Round(100 * DeltaY2 / DeltaX2, 1);
                                int Cat2 = find_category(dt_slope_ranges, Slope2);

                                string plus = "";
                                if (Slope2 > 0) plus = "+";
                                if (Cat2 != -1)
                                {
                                    MText mt1 = new MText();
                                    mt1.Contents = plus + Functions.Get_String_Rounded(Slope2, 1) + "% - Rge " + Cat2.ToString() + "\r\nd=" + Functions.Get_String_Rounded(Dist, 1);
                                    mt1.Attachment = AttachmentPoint.TopCenter;
                                    mt1.Location = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                                    mt1.TextHeight = Texth;
                                    mt1.Rotation = Rot1;
                                    mt1.Layer = nume_layer;
                                    mt1.ColorIndex = color_index(Cat2);

                                    BTrecord.AppendEntity(mt1);
                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                    List_txt.Insert(index1, mt1.ObjectId);
                                }
                                else
                                {
                                    List_txt.Insert(index1, ObjectId.Null);
                                }

                                Update_blocks_slopeID(Trans1, BTrecord, Poly_avg1);

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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();

        }



        private void Button_add_div_berm_for_entire_profile_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);

            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;

            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            if (dt_terrain == null || dt_terrain.Rows.Count == 0)
            {
                MessageBox.Show("No terrains loaded. Load the terrain table first");
                return;
            }

            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    this.WindowState = FormWindowState.Minimized;



                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly1;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_poly1.MessageForAdding = "\nselect optimized slope polyline:";
                    Prompt_poly1.SingleOnly = true;


                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                    Prompt_poly2.SingleOnly = true;

                    Rezultat_poly1 = null;
                    Rezultat_poly2 = null;

                    if (slope_id == ObjectId.Null || profile_id == ObjectId.Null)
                    {
                        Editor1.SetImpliedSelection(Empty_array);
                        Rezultat_poly1 = ThisDrawing.Editor.GetSelection(Prompt_poly1);


                        if (Rezultat_poly1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();

                            panel_poly_profile.BackColor = Color.Red;
                            panel_poly_slope.BackColor = Color.Red;
                            this.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Editor1.SetImpliedSelection(Empty_array);
                        Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                        if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();

                            panel_poly_profile.BackColor = Color.Red;
                            panel_poly_slope.BackColor = Color.Red;
                            this.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        panel_poly_profile.BackColor = Color.Green;
                        panel_poly_slope.BackColor = Color.Green;
                    }



                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        if (slope_id == ObjectId.Null)
                        {
                            slope_id = Rezultat_poly1.Value[0].ObjectId;
                        }

                        Entity Ent1 = Trans1.GetObject(slope_id, OpenMode.ForRead) as Entity;
                        if ((Ent1 is Polyline) == false)
                        {
                            MessageBox.Show("the optimized slope is not a polyline\r\n" + Ent1.GetType().ToString() + "\r\nOperation aborted");
                            this.WindowState = FormWindowState.Normal;
                            slope_id = ObjectId.Null;
                            profile_id = ObjectId.Null;
                            set_enable_true();
                            panel_poly_slope.BackColor = Color.Red;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        if (profile_id == ObjectId.Null)
                        {
                            profile_id = Rezultat_poly2.Value[0].ObjectId;
                        }

                        Entity Ent2 = Trans1.GetObject(profile_id, OpenMode.ForRead) as Entity;
                        if ((Ent2 is Polyline) == false)
                        {
                            MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent2.GetType().ToString() + "\r\nOperation aborted");
                            this.WindowState = FormWindowState.Normal;
                            slope_id = ObjectId.Null;
                            profile_id = ObjectId.Null;
                            set_enable_true();
                            panel_poly_profile.BackColor = Color.Red;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Slope_poly = Ent1 as Polyline;
                        Prof_poly = Ent2 as Polyline;
                        string ln = "DivBerm";

                        Functions.Creaza_layer(ln, 1, true);

                        if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                        {
                            for (int k = dt_terrain.Rows.Count - 1; k > 0; --k)
                            {
                                if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value && dt_terrain.Rows[k - 1][0] != DBNull.Value && dt_terrain.Rows[k - 1][1] != DBNull.Value)
                                {
                                    int cat1 = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                    double sta1 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                    int cat2 = Convert.ToInt32(dt_terrain.Rows[k - 1][2]);
                                    if (cat1 == cat2)
                                    {
                                        dt_terrain.Rows[k - 1][1] = sta1;
                                        dt_terrain.Rows[k].Delete();
                                    }
                                }
                            }
                        }


                        for (int i = 0; i < Slope_poly.NumberOfVertices - 1; ++i)
                        {
                            Point3d pt1 = Slope_poly.GetPointAtParameter(i);
                            Point3d pt2 = Slope_poly.GetPointAtParameter(i + 1);
                            double sta01 = sta0 + pt1.X - Prof_poly.StartPoint.X;
                            double sta02 = sta0 + pt2.X - Prof_poly.StartPoint.X;

                            double slope = calc_slope(pt1, pt2);

                            if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                            {
                                double min_dist = 1000000;
                                if (Functions.IsNumeric(textBox_min_berm.Text) == true)
                                {
                                    min_dist = Convert.ToDouble(textBox_min_berm.Text);
                                }

                                double cumul_dist = 0;
                                for (int k = 0; k < dt_terrain.Rows.Count; ++k)
                                {
                                    if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value)
                                    {
                                        double t1 = Convert.ToDouble(dt_terrain.Rows[k][0]);
                                        double t2 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                        int category = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                        if (sta01 >= t1 && sta02 <= t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = pt1;
                                            xline1.SecondPoint = new Point3d(pt1.X, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = pt2;
                                            xline2.SecondPoint = new Point3d(pt2.X, pt2.Y + 10, pt2.Z);
                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;






                                                double dist = sta2 - sta1;
                                                cumul_dist = dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + 1, category, pt1, ln);


                                            }
                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;
                                        }
                                        else if (sta01 > t1 && sta02 > t2 && sta01 < t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = pt1;
                                            xline1.SecondPoint = new Point3d(pt1.X, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = new Point3d(pt1.X + t2 - sta01, pt1.Y, pt1.Z);
                                            xline2.SecondPoint = new Point3d(pt1.X + t2 - sta01, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + 1, category, pt1, ln);
                                            }

                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;


                                        }
                                        else if (sta01 < t1 && sta02 < t2 && sta02 > t1)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = new Point3d(pt1.X + t1 - sta01, pt1.Y, pt1.Z);
                                            xline1.SecondPoint = new Point3d(pt1.X + t1 - sta01, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = pt2;
                                            xline2.SecondPoint = new Point3d(pt2.X, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + 1, category, pt1, ln);
                                            }

                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;

                                        }
                                        else if (sta01 <= t1 && sta02 >= t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = new Point3d(pt1.X + t1 - sta01, pt1.Y, pt1.Z);
                                            xline1.SecondPoint = new Point3d(pt1.X + t1 - sta01, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = new Point3d(pt1.X + t2 - sta01, pt1.Y, pt1.Z);
                                            xline2.SecondPoint = new Point3d(pt1.X + t2 - sta01, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + 1, category, pt1, ln);
                                            }



                                        }
                                        else if (sta02 < t1 && sta02 < t2)
                                        {
                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;
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
            set_enable_true();
            this.WindowState = FormWindowState.Normal;


        }
        void insert_div_berm_block(Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing, BlockTableRecord BTrecord,
            Point3d colint1, Point3d colint2, double sta0, double cumul_dist, double min_dist, double slope, int slope_index, int category, Point3d pt1, string ln)
        {


            double sta1 = sta0 + colint1.X - Prof_poly.StartPoint.X;
            double sta2 = sta0 + colint2.X - Prof_poly.StartPoint.X;
            double dist = sta2 - sta1;
            string note1 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + "Hr" + DateTime.Now.Minute.ToString() + "Min_by" + Environment.UserName.ToUpper();

            if (cumul_dist >= min_dist)
            {
                double spacing = get_spacing(Math.Abs(slope), category);

                if (spacing > 0)
                {
                    double no1 = dist / spacing;
                    int no_of_spaces = Convert.ToInt32(Math.Floor(no1));

                    if (no_of_spaces == 0)
                    {
                        Xline xline3 = new Xline();
                        xline3.BasePoint = new Point3d(pt1.X + dist / 2, pt1.Y, pt1.Z);
                        xline3.SecondPoint = new Point3d(pt1.X + dist / 2, pt1.Y + 10, pt1.Z);
                        Point3dCollection colint3 = Functions.Intersect_on_both_operands(Prof_poly, xline3);
                        if (colint3.Count > 0)
                        {
                            System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                            System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                            double staX = sta0 + colint3[0].X - Prof_poly.StartPoint.X;


                            if (dt_prof != null && dt_prof.Rows.Count > 0)
                            {
                                double param1 = Prof_poly.GetParameterAtPoint(Prof_poly.GetClosestPointTo(colint3[0], Vector3d.ZAxis, false));
                                int idx0 = Convert.ToInt32(Math.Floor(param1));
                                double dif1 = Prof_poly.GetPointAtParameter(param1).X - Prof_poly.GetPointAtParameter(idx0).X;

                                if (dt_prof.Rows.Count >= idx0 + 1)
                                {
                                    if (dt_prof.Rows[idx0][Col_sta] != DBNull.Value)
                                    {
                                        double sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta]);
                                        if (dt_prof.Rows[idx0][Col_sta_eq] != DBNull.Value)
                                        {
                                            sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta_eq]);
                                        }
                                        staX = sta_start + dif1;
                                    }
                                }

                            }

                            string stax_string = Functions.Get_chainage_from_double(staX, "m", 1);
                            string slope_string = "AVG SLOPE = " + Functions.Get_String_Rounded(slope, 1) + "%";
                            string spacing_string = "SPACING = " + Functions.Get_String_Rounded(spacing, 1);
                            col_atr.Add("STA");
                            col_val.Add(stax_string);
                            col_atr.Add("SLOPE");
                            col_val.Add(slope_string);
                            col_atr.Add("SLOPEID");
                            col_val.Add(Convert.ToString(slope_index));
                            col_atr.Add("SPACING");
                            col_val.Add(spacing_string);
                            col_atr.Add("NO");
                            col_val.Add("1 of 1");
                            col_atr.Add("DITCHPLUG");
                            col_val.Add("NO");

                            col_atr.Add("TERRAIN");
                            if (category == 1)
                            {
                                col_val.Add("Fine Sand");
                            }
                            if (category == 2)
                            {
                                col_val.Add("Clay");
                            }
                            if (category == 3)
                            {
                                col_val.Add("Gravel/Bedrock");
                            }

                            col_atr.Add("NOTE");
                            col_val.Add(note1);
                            col_atr.Add("NOTES");
                            col_val.Add(note1);


                            BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", "DB", colint3[0], 0.1, 0, ln, col_atr, col_val);
                        }
                    }
                    else
                    {
                        double dif = dist - (no_of_spaces * spacing);

                        for (int j = 0; j <= no_of_spaces; ++j)
                        {
                            Xline xline3 = new Xline();
                            xline3.BasePoint = new Point3d(pt1.X + j * spacing + dif / 2, pt1.Y, pt1.Z);
                            xline3.SecondPoint = new Point3d(pt1.X + j * spacing + dif / 2, pt1.Y + 10, pt1.Z);
                            Point3dCollection colint3 = Functions.Intersect_on_both_operands(Prof_poly, xline3);
                            if (colint3.Count > 0)
                            {
                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                double staX = sta0 + colint3[0].X - Prof_poly.StartPoint.X;


                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                {
                                    double param1 = Prof_poly.GetParameterAtPoint(Prof_poly.GetClosestPointTo(colint3[0], Vector3d.ZAxis, false));
                                    int idx0 = Convert.ToInt32(Math.Floor(param1));
                                    double dif1 = Prof_poly.GetPointAtParameter(param1).X - Prof_poly.GetPointAtParameter(idx0).X;

                                    if (dt_prof.Rows.Count >= idx0 + 1)
                                    {
                                        if (dt_prof.Rows[idx0][Col_sta] != DBNull.Value)
                                        {
                                            double sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta]);
                                            if (dt_prof.Rows[idx0][Col_sta_eq] != DBNull.Value)
                                            {
                                                sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta_eq]);
                                            }
                                            staX = sta_start + dif1;
                                        }
                                    }

                                }

                                string staX_string = Functions.Get_chainage_from_double(staX, "m", 1);
                                string slope_string = "AVG SLOPE = " + Functions.Get_String_Rounded(slope, 1) + "%";
                                string spacing_string = "SPACING = " + Functions.Get_String_Rounded(spacing, 1);
                                string no_string = Convert.ToString(j + 1) + " of " + Convert.ToString(no_of_spaces + 1);
                                col_atr.Add("STA");
                                col_val.Add(staX_string);
                                col_atr.Add("SLOPE");
                                col_val.Add(slope_string);
                                col_atr.Add("SLOPEID");
                                col_val.Add(Convert.ToString(slope_index));
                                col_atr.Add("SPACING");
                                col_val.Add(spacing_string);
                                col_atr.Add("NO");
                                col_val.Add(no_string);
                                col_atr.Add("DITCHPLUG");
                                col_val.Add("NO");

                                col_atr.Add("TERRAIN");
                                if (category == 1)
                                {
                                    col_val.Add("Fine Sand");
                                }
                                if (category == 2)
                                {
                                    col_val.Add("Clay");
                                }
                                if (category == 3)
                                {
                                    col_val.Add("Gravel/Bedrock");
                                }

                                col_atr.Add("NOTE");
                                col_val.Add(note1);
                                col_atr.Add("NOTES");
                                col_val.Add(note1);

                                BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", "DB", colint3[0], 0.1, 0, ln, col_atr, col_val);
                            }
                        }
                    }
                }


            }
            cumul_dist = 0;

        }


        private double get_spacing(double slope, int category)
        {
            double spacing = 0;
            double spcc = -1;

            if (Functions.IsNumeric(textBox_ss1.Text) == true && Functions.IsNumeric(textBox_se1.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss1.Text);
                double slope2 = Convert.ToDouble(textBox_se1.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 1);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss2.Text) == true && Functions.IsNumeric(textBox_se2.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss2.Text);
                double slope2 = Convert.ToDouble(textBox_se2.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 2);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss3.Text) == true && Functions.IsNumeric(textBox_se3.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss3.Text);
                double slope2 = Convert.ToDouble(textBox_se3.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 3);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss4.Text) == true && Functions.IsNumeric(textBox_se4.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss4.Text);
                double slope2 = Convert.ToDouble(textBox_se4.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 4);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss5.Text) == true && Functions.IsNumeric(textBox_se5.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss5.Text);
                double slope2 = Convert.ToDouble(textBox_se5.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 5);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss6.Text) == true && Functions.IsNumeric(textBox_se6.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss6.Text);
                double slope2 = Convert.ToDouble(textBox_se6.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 6);
            }

            if (spcc == -1 && Functions.IsNumeric(textBox_ss7.Text) == true && Functions.IsNumeric(textBox_se7.Text) == true)
            {
                double slope1 = Convert.ToDouble(textBox_ss7.Text);
                double slope2 = Convert.ToDouble(textBox_se7.Text);
                spcc = get_spacing_2_slopes(slope, category, slope1, slope2, 7);
            }

            if (spcc >= 0) spacing = spcc;

            return spacing;
        }


        private double get_spacing_2_slopes(double slope, int category, double slope1, double slope2, int row)
        {
            double spcc = -1;
            if (Math.Abs(slope) >= slope1 && Math.Abs(slope) < slope2)
            {
                string formula1 = "";

                if (row == 1)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing1A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing1B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing1C.Text;
                    }
                }

                if (row == 2)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing2A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing2B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing2C.Text;
                    }
                }

                if (row == 3)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing3A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing3B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing3C.Text;
                    }
                }

                if (row == 4)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing4A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing4B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing4C.Text;
                    }
                }

                if (row == 5)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing5A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing5B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing5C.Text;
                    }
                }

                if (row == 6)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing6A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing6B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing6C.Text;
                    }
                }

                if (row == 7)
                {
                    if (category == 1)
                    {
                        formula1 = textBox_spacing7A.Text;
                    }
                    if (category == 2)
                    {
                        formula1 = textBox_spacing7B.Text;
                    }
                    if (category == 3)
                    {
                        formula1 = textBox_spacing7C.Text;
                    }
                }

                if (Functions.IsNumeric(formula1) == true)
                {
                    spcc = Convert.ToDouble(formula1);
                }
                else
                {
                    if (formula1.Replace(" ", "").Length > 0)
                    {
                        string continut = formula1.Replace(" ", "");
                        int index1 = 0;
                        if (continut.Contains("/") == true)
                        {
                            index1 = continut.IndexOf("/");
                        }
                        int index2 = 0;
                        if (continut.Contains("*") == true)
                        {
                            index2 = continut.IndexOf("*");
                        }

                        if (index1 > 0 && index2 > 0)
                        {
                            if (index1 > index2)
                            {
                                string numar1 = continut.Substring(0, index2);
                                string numar2 = continut.Substring(index2 + 1, index1 - index2 - 1);
                                string numar3 = continut.Substring(index1 + 1, continut.Length - index1 - 1);
                                spcc = get_spacing_3_numbers(numar1, numar2, numar3, slope);
                            }
                            if (index1 < index2)
                            {
                                string numar1 = continut.Substring(0, index1);
                                string numar2 = continut.Substring(index1 + 1, index2 - index1 - 1);
                                string numar3 = continut.Substring(index2 + 1, continut.Length - index2 - 1);
                                spcc = get_spacing_3_numbers(numar1, numar2, numar3, slope);
                            }


                        }
                        if (index1 > 0 && index2 == 0)
                        {
                            string numar1 = continut.Substring(0, index1);
                            string numar2 = continut.Substring(index1 + 1, continut.Length - index1 - 1);
                            spcc = get_spacing_2_numbers_division(numar1, numar2, slope);
                        }
                        if (index1 == 0 && index2 > 0)
                        {
                            string numar1 = continut.Substring(0, index2);
                            string numar2 = continut.Substring(index2 + 1, continut.Length - index2 - 1);
                            spcc = get_spacing_2_numbers_multiply(numar1, numar2, slope);
                        }

                    }
                }



            }

            return spcc;
        }
        private double get_spacing_3_numbers(string numar1, string numar2, string numar3, double slope)
        {
            double spacing = 0;
            if (numar3 == "SLOPE" && Functions.IsNumeric(numar1) == true && Functions.IsNumeric(numar2) == true)
            {
                spacing = Math.Round(Convert.ToDouble(numar1) * Convert.ToDouble(numar2) / slope, 1);
            }
            if (numar1 == "SLOPE" && Functions.IsNumeric(numar2) == true && Functions.IsNumeric(numar3) == true)
            {
                spacing = Math.Round(slope * Convert.ToDouble(numar2) / Convert.ToDouble(numar3), 1);
            }
            if (numar2 == "SLOPE" && Functions.IsNumeric(numar1) == true && Functions.IsNumeric(numar3) == true)
            {
                spacing = Math.Round(slope * Convert.ToDouble(numar1) / Convert.ToDouble(numar3), 1);
            }

            return spacing;
        }

        private double get_spacing_2_numbers_division(string numar1, string numar2, double slope)
        {
            double spacing = 0;
            if (numar2 == "SLOPE" && Functions.IsNumeric(numar1) == true)
            {
                spacing = Math.Round(Convert.ToDouble(numar1) / slope, 1);
            }
            if (numar1 == "SLOPE" && Functions.IsNumeric(numar2) == true)
            {
                spacing = Math.Round(slope / Convert.ToDouble(numar2), 1);
            }
            return spacing;
        }

        private double get_spacing_2_numbers_multiply(string numar1, string numar2, double slope)
        {
            double spacing = 0;
            if (numar2 == "SLOPE" && Functions.IsNumeric(numar1) == true)
            {
                spacing = Math.Round(Convert.ToDouble(numar1) * slope, 1);
            }
            if (numar1 == "SLOPE" && Functions.IsNumeric(numar2) == true)
            {
                spacing = Math.Round(slope * Convert.ToDouble(numar2), 1);
            }
            return spacing;
        }

        public System.Data.DataTable build_soils_data_table_from_excel_based_on_columns_with_type_check(System.Data.DataTable dt1, Microsoft.Office.Interop.Excel.Worksheet W1,
            int start_row, int end_row, string col1, string colxl1, string col2, string colxl2, string col3, string colxl3)

        {
            if (W1 == null) return dt1;
            if (end_row - start_row < 0) return dt1;

            object[,] values1 = new object[end_row - start_row + 1, 1];
            object[,] values2 = new object[end_row - start_row + 1, 1];
            object[,] values3 = new object[end_row - start_row + 1, 1];

            #region 1
            if (colxl1 != "")
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[colxl1 + start_row.ToString() + ":" + colxl1 + end_row.ToString()];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    dt1.Rows.Add();

                    if (Valoare1 == null) Valoare1 = DBNull.Value;

                    if (Valoare1 != null && dt1.Columns[col1].GetType() == typeof(double))
                    {
                        Valoare1 = Valoare1.ToString().Replace("+", "");
                        if (Functions.IsNumeric(Valoare1.ToString()) == true)
                        {
                            Valoare1 = Convert.ToDouble(Valoare1);
                        }
                        else
                        {
                            Valoare1 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col1] = Valoare1;
                }
            }
            #endregion

            #region 2
            if (colxl2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colxl2 + start_row.ToString() + ":" + colxl2 + end_row.ToString()];
                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (colxl1 == "") dt1.Rows.Add();

                    if (Valoare2 == null) Valoare2 = DBNull.Value;

                    if (Valoare2 != null && dt1.Columns[col2].GetType() == typeof(double))
                    {
                        Valoare2 = Valoare2.ToString().Replace("+", "");
                        if (Functions.IsNumeric(Valoare2.ToString()) == true)
                        {
                            Valoare2 = Convert.ToDouble(Valoare2);
                        }
                        else
                        {
                            Valoare2 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col2] = Valoare2;
                }
            }
            #endregion

            #region 3
            if (colxl3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colxl3 + start_row.ToString() + ":" + colxl3 + end_row.ToString()];
                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (colxl1 == "" && colxl2 == "") dt1.Rows.Add();

                    if (Valoare3 == null) Valoare3 = DBNull.Value;

                    if (Valoare3 != null)
                    {
                        string Val3 = Valoare3.ToString();
                        if (Val3.ToLower().Contains("organic") == true || Val3.ToLower().Contains("sand") == true)
                        {
                            dt1.Rows[i - 1][col3] = 1;
                        }
                        else if (Val3.ToLower().Contains("clay") == true)
                        {
                            dt1.Rows[i - 1][col3] = 2;
                        }
                        else if (Val3.ToLower().Contains("gravel") == true || Val3.ToLower().Contains("rock") == true)
                        {
                            dt1.Rows[i - 1][col3] = 3;
                        }
                        else
                        {
                            dt1.Rows[i - 1][col3] = 1;
                        }

                    }
                    else
                    {
                        dt1.Rows[i - 1][col3] = 1;
                    }

                }
            }
            #endregion

            return dt1;
        }

        public System.Data.DataTable build_terrain_data_table_from_excel(System.Data.DataTable dt1, Microsoft.Office.Interop.Excel.Worksheet W1,
          int start_row, int end_row, string col1, string colxl1, string col2, string colxl2, string col3, string colxl3, string col4, string colxl4, string col5, string colxl5)

        {
            if (W1 == null) return dt1;
            if (end_row - start_row < 0) return dt1;

            object[,] values1 = new object[end_row - start_row + 1, 1];
            object[,] values2 = new object[end_row - start_row + 1, 1];
            object[,] values3 = new object[end_row - start_row + 1, 1];
            object[,] values4 = new object[end_row - start_row + 1, 1];
            object[,] values5 = new object[end_row - start_row + 1, 1];

            #region 1
            if (colxl1 != "")
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[colxl1 + start_row.ToString() + ":" + colxl1 + end_row.ToString()];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    dt1.Rows.Add();

                    if (Valoare1 == null) Valoare1 = DBNull.Value;

                    if (Valoare1 != null && dt1.Columns[col1].GetType() == typeof(double))
                    {
                        Valoare1 = Valoare1.ToString().Replace("+", "");
                        if (Functions.IsNumeric(Valoare1.ToString()) == true)
                        {
                            Valoare1 = Convert.ToDouble(Valoare1);
                        }
                        else
                        {
                            Valoare1 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col1] = Valoare1;
                }
            }
            #endregion

            #region 2
            if (colxl2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colxl2 + start_row.ToString() + ":" + colxl2 + end_row.ToString()];
                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (colxl1 == "") dt1.Rows.Add();

                    if (Valoare2 == null) Valoare2 = DBNull.Value;

                    if (Valoare2 != null && dt1.Columns[col2].GetType() == typeof(double))
                    {
                        Valoare2 = Valoare2.ToString().Replace("+", "");
                        if (Functions.IsNumeric(Valoare2.ToString()) == true)
                        {
                            Valoare2 = Convert.ToDouble(Valoare2);
                        }
                        else
                        {
                            Valoare2 = DBNull.Value;
                        }
                    }

                    dt1.Rows[i - 1][col2] = Valoare2;
                }
            }
            #endregion

            #region 3
            if (colxl3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colxl3 + start_row.ToString() + ":" + colxl3 + end_row.ToString()];
                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (colxl1 == "" && colxl2 == "") dt1.Rows.Add();

                    if (Valoare3 == null) Valoare3 = "-";
                    dt1.Rows[i - 1][col3] = Convert.ToString(Valoare3);
                }
            }
            #endregion

            #region 4
            if (colxl4 != "")
            {
                Microsoft.Office.Interop.Excel.Range range4 = W1.Range[colxl4 + start_row.ToString() + ":" + colxl4 + end_row.ToString()];
                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (colxl1 == "" && colxl2 == "" && colxl3 == "") dt1.Rows.Add();

                    if (Valoare4 == null) Valoare4 = "-";
                    dt1.Rows[i - 1][col4] = Convert.ToString(Valoare4);
                }
            }
            #endregion

            #region 5
            if (colxl5 != "")
            {
                Microsoft.Office.Interop.Excel.Range range5 = W1.Range[colxl5 + start_row.ToString() + ":" + colxl5 + end_row.ToString()];
                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (colxl1 == "" && colxl2 == "" && colxl3 == "" && colxl4 == "") dt1.Rows.Add();

                    if (Valoare5 == null) Valoare5 = "-";
                    dt1.Rows[i - 1][col5] = Convert.ToString(Valoare5);
                }
            }
            #endregion




            return dt1;
        }

        private void Button_load_soils_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }

            if (Functions.IsNumeric(textBox_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            if (comboBox_ws1.Text == "")
            {
                MessageBox.Show("No terrain info has been loaded because you did not specified the excel spreadsheet!");
                return;
            }

            try
            {
                if (comboBox_ws1.Text != "")
                {
                    string string1 = comboBox_ws1.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            set_enable_false();
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null)
                            {
                                dt_terrain = new System.Data.DataTable();
                                dt_terrain.Columns.Add("sta1", typeof(double));
                                dt_terrain.Columns.Add("sta2", typeof(double));
                                dt_terrain.Columns.Add("soil_category", typeof(int));

                                dt_terrain = build_soils_data_table_from_excel_based_on_columns_with_type_check(dt_terrain, W1, start1, end1, "sta1", textBox_sta1.Text, "sta2", textBox_sta2.Text, "soil_category", textBox_soil.Text);

                                dt_terrain = simplify_dt_terain(dt_terrain, "soil_category");

                                MessageBox.Show("done", "divberm", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        }

        private void Button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
            {
                string item1 = Convert.ToString(comboBox_ws1.Items[i]);
                if (item1.ToLower().Contains("[materials]") == true)
                {
                    comboBox_ws1.SelectedIndex = i;
                    i = comboBox_ws1.Items.Count;
                }
            }
        }

        private void Label_column_mapping_Click(object sender, EventArgs e)
        {
            if (panel_terrain.Visible == true)
            {
                panel_terrain.Visible = false;
            }
            else
            {
                panel_terrain.Visible = true;
            }
        }

        private System.Data.DataTable simplify_dt_terain(System.Data.DataTable dt1 ,string soil_column_name )
        {
            if (dt1 != null && dt1.Rows.Count > 0)
            {
                for (int i = dt1.Rows.Count - 1; i > 0; --i)
                {
                    if (dt1.Rows[i]["sta2"] != DBNull.Value)
                    {
                        double sta2 = Convert.ToDouble(dt1.Rows[i]["sta2"]);
                        if (dt1.Rows[i][soil_column_name] != DBNull.Value)
                        {
                            string val1 = Convert.ToString(dt1.Rows[i][soil_column_name]);
                            for (int j = i - 1; j >= 0; --j)
                            {

                                if (dt1.Rows[j][soil_column_name] != DBNull.Value)
                                {
                                    string val2 = Convert.ToString(dt1.Rows[j][soil_column_name]);
                                    if (val1 == val2)
                                    {
                                        dt1.Rows[j]["sta2"] = sta2;
                                        dt1.Rows[i].Delete();
                                    }
                                    j = -1;
                                }
                            }
                        }


                    }
                }
            }

            return dt1;
        }

        private void Button_insert_terrain_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }

            if (Functions.IsNumeric(textBox_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }
            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);
            try
            {


                if (comboBox_ws1.Text != "")
                {
                    string string1 = comboBox_ws1.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            set_enable_false();
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null)
                            {
                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("sta1", typeof(double));
                                dt1.Columns.Add("sta2", typeof(double));
                                dt1.Columns.Add("line1", typeof(string));
                                dt1.Columns.Add("line2", typeof(string));
                                dt1.Columns.Add("line3", typeof(string));

                                dt1 = build_terrain_data_table_from_excel(dt1, W1, start1, end1, "sta1", textBox_sta1.Text, "sta2", textBox_sta2.Text, "line2", textBox_soil.Text, "line1", textBox_golder.Text, "line3", textBox_model.Text);
                                if (dt1.Rows.Count > 0)
                                {

                                    dt1 = simplify_dt_terain(dt1, "line2");

                                    ObjectId[] Empty_array = null;
                                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                                    Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                    {
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                        Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                                        Prompt_poly2.SingleOnly = true;
                                        this.WindowState = FormWindowState.Minimized;
                                        Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                                        if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            this.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {
                                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                            Entity Ent2 = Trans1.GetObject(Rezultat_poly2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                            if ((Ent2 is Polyline) == false)
                                            {
                                                MessageBox.Show("the polyline profile is not a polyline\r\n is a " + Ent2.GetType().ToString() + "\r\nOperation aborted");
                                                this.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }
                                            Polyline poly1 = Ent2 as Polyline;
                                            double ymin = -1000000;
                                            double ymax = 1000000;

                                            for (int i = 0; i < poly1.NumberOfVertices - 1; ++i)
                                            {
                                                double y = poly1.GetPointAtParameter(i).Y;
                                                if (i == 0)
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

                                            string ter_ln = "_terrain_blocks";
                                            string block_name = "Terrainlr";
                                            Functions.Creaza_layer(ter_ln, 2, true);

                                            double deltaY = 0;

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                double sta1 = Convert.ToDouble(dt1.Rows[i][0]);
                                                double sta2 = Convert.ToDouble(dt1.Rows[i][1]);
                                                string line1 = Convert.ToString(dt1.Rows[i][2]);
                                                string line2 = Convert.ToString(dt1.Rows[i][3]);
                                                string line3 = Convert.ToString(dt1.Rows[i][4]);


                                                double diferenta1 = 0;
                                                double diferenta2 = 0;



                                                int index_sta1 = -1;
                                                int index_sta2 = -1;
                                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][Col_sta] != DBNull.Value && dt_prof.Rows[j + 1][Col_sta] != DBNull.Value)
                                                        {
                                                            double sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta]);
                                                            if (dt_prof.Rows[j][Col_sta_eq] != DBNull.Value)
                                                            {
                                                                sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta_eq]);
                                                            }
                                                            double sta_end = Convert.ToDouble(dt_prof.Rows[j + 1][Col_sta]);

                                                            if (index_sta1 == -1 && sta1 >= sta_start && sta1 <= sta_end)
                                                            {
                                                                index_sta1 = j;
                                                                diferenta1 = sta1 - sta_start;
                                                            }
                                                            if (index_sta2 == -1 && sta2 >= sta_start && sta2 <= sta_end)
                                                            {
                                                                index_sta2 = j;
                                                                diferenta2 = sta2 - sta_start;
                                                            }
                                                            if (index_sta1 != -1 && index_sta2 != -1)
                                                            {
                                                                j = dt_prof.Rows.Count;
                                                            }

                                                        }
                                                    }
                                                }



                                                double x1 = poly1.StartPoint.X + sta0 + sta1;

                                                if (index_sta1 >= 0)
                                                {
                                                    x1 = poly1.GetPoint2dAt(index_sta1).X + diferenta1;
                                                }

                                                Xline xline1 = new Xline();
                                                xline1.BasePoint = new Point3d(x1, 0, poly1.Elevation);
                                                xline1.SecondPoint = new Point3d(x1, 10, poly1.Elevation);

                                                Point3dCollection col1 = Functions.Intersect_on_both_operands(poly1, xline1);
                                                Point3d inspt1 = new Point3d();
                                                if (col1.Count > 0)
                                                {
                                                    inspt1 = col1[0];
                                                }
                                                else
                                                {
                                                    inspt1 = new Point3d(x1, (ymin + ymax) / 2, 0);
                                                }


                                                double x2 = poly1.StartPoint.X + sta0 + sta2;


                                                if (index_sta2 >= 0)
                                                {
                                                    double x222 = poly1.GetPoint2dAt(index_sta2).X;

                                                    x2 = poly1.GetPoint2dAt(index_sta2).X + diferenta2;
                                                }

                                                Xline xline2 = new Xline();
                                                xline2.BasePoint = new Point3d(x2, 0, poly1.Elevation);
                                                xline2.SecondPoint = new Point3d(x2, 10, poly1.Elevation);

                                                Point3dCollection col2 = Functions.Intersect_on_both_operands(poly1, xline2);
                                                Point3d inspt2 = new Point3d();
                                                if (col2.Count > 0)
                                                {
                                                    inspt2 = col2[0];
                                                }
                                                else
                                                {
                                                    inspt2 = new Point3d(x2, (ymin + ymax) / 2, 0);
                                                }

                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                string display_sta_string1 = Functions.Get_chainage_from_double(sta1, "m", 1);
                                                string display_sta_string2 = Functions.Get_chainage_from_double(sta2, "m", 1);

                                                col_atr.Add("DESCR");
                                                col_val.Add(line1 + "\r\n" + line2 + "\r\n" + line3);

                                                col_atr.Add("STA1");
                                                col_val.Add(display_sta_string1);

                                                col_atr.Add("STA2");
                                                col_val.Add(display_sta_string2);


                                                BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block_name, inspt1, 1, 0, ter_ln, col_atr, col_val);
                                                double exist_dist2 = Functions.Get_Param_Value_block(block1, "Distance2");
                                                double exist_dist3 = Functions.Get_Param_Value_block(block1, "Distance3");

                                                if (sta2 < sta1)
                                                {
                                                    MessageBox.Show("Sta2 " + sta2.ToString() + " is smaller than \r\nSta1 " + sta1.ToString());
                                                }

                                                double diference = 0;
                                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][Col_sta] != DBNull.Value && dt_prof.Rows[j][Col_sta_eq] != DBNull.Value)
                                                        {
                                                            double sta_back = Convert.ToDouble(dt_prof.Rows[j][Col_sta]);

                                                            double sta_ahead = Convert.ToDouble(dt_prof.Rows[j][Col_sta_eq]);
                                                            if (sta2 > sta_ahead && sta1 < sta_back)
                                                            {
                                                                diference = diference + sta_ahead - sta_back;
                                                            }

                                                        }
                                                    }
                                                }

                                                double stretch1 = sta2 - sta1;

                                                Functions.Stretch_block(block1, "Distance1", stretch1 - diference);
                                                Functions.Stretch_block(block1, "Distance2", exist_dist2 + (inspt1.Y - inspt2.Y));
                                                Functions.Stretch_block(block1, "Distance3", exist_dist3 + deltaY);
                                              
                                                if (deltaY == 0)
                                                {
                                                    deltaY = 200;
                                                }
                                                else if (deltaY == 200)
                                                {
                                                    deltaY = -200;
                                                }
                                                else if (deltaY == -200)
                                                {
                                                    deltaY = 0;
                                                }

                                            }
                                            Trans1.Commit();
                                        }

                                    }
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
            ThisDrawing.Editor.WriteMessage("\nCommand:");
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }

        private void Button_place_div_berms_for_a_profile_portion_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);

            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            if (dt_terrain == null || dt_terrain.Rows.Count == 0)
            {
                MessageBox.Show("No terrains loaded. Load the terrain table first");
                return;
            }
            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    this.WindowState = FormWindowState.Minimized;

                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly1;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_poly1.MessageForAdding = "\nselect optimized slope polyline:";
                    Prompt_poly1.SingleOnly = true;

                    Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                    Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                    Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                    Prompt_poly2.SingleOnly = true;

                    Rezultat_poly1 = null;
                    Rezultat_poly2 = null;

                    if (slope_id == ObjectId.Null || profile_id == ObjectId.Null)
                    {
                        Editor1.SetImpliedSelection(Empty_array);
                        Rezultat_poly1 = ThisDrawing.Editor.GetSelection(Prompt_poly1);


                        if (Rezultat_poly1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();

                            panel_poly_profile.BackColor = Color.Red;
                            panel_poly_slope.BackColor = Color.Red;
                            this.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Editor1.SetImpliedSelection(Empty_array);
                        Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                        if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();

                            panel_poly_profile.BackColor = Color.Red;
                            panel_poly_slope.BackColor = Color.Red;
                            this.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        panel_poly_profile.BackColor = Color.Green;
                        panel_poly_slope.BackColor = Color.Green;
                    }


                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        if (slope_id == ObjectId.Null)
                        {
                            slope_id = Rezultat_poly1.Value[0].ObjectId;
                        }

                        if (profile_id == ObjectId.Null)
                        {
                            profile_id = Rezultat_poly2.Value[0].ObjectId;
                        }

                        Entity Ent1 = Trans1.GetObject(slope_id, OpenMode.ForRead) as Entity;
                        if ((Ent1 is Polyline) == false)
                        {
                            MessageBox.Show("the optimized slope is not a polyline\r\n" + Ent1.GetType().ToString() + "\r\nOperation aborted");
                            this.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            panel_poly_slope.BackColor = Color.Red;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Entity Ent2 = Trans1.GetObject(profile_id, OpenMode.ForRead) as Entity;
                        if ((Ent2 is Polyline) == false)
                        {
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;
                            panel_poly_profile.BackColor = Color.Red;
                            MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent2.GetType().ToString() + "\r\nOperation aborted");
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Slope_poly = Ent1 as Polyline;
                        Prof_poly = Ent2 as Polyline;

                        string ln = "DivBerm";

                        Functions.Creaza_layer(ln, 1, true);

                        if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                        {
                            for (int k = dt_terrain.Rows.Count - 1; k > 0; --k)
                            {
                                if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value && dt_terrain.Rows[k - 1][0] != DBNull.Value && dt_terrain.Rows[k - 1][1] != DBNull.Value)
                                {
                                    int cat1 = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                    double sta1 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                    int cat2 = Convert.ToInt32(dt_terrain.Rows[k - 1][2]);
                                    if (cat1 == cat2)
                                    {
                                        dt_terrain.Rows[k - 1][1] = sta1;
                                        dt_terrain.Rows[k].Delete();
                                    }
                                }
                            }
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                        PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the start point");
                        PP5.AllowNone = false;
                        Point_res5 = Editor1.GetPoint(PP5);

                        if (Point_res5.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;

                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res6;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP6;
                        PP6 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the end point");
                        PP6.AllowNone = false;
                        PP6.UseBasePoint = true;
                        PP6.BasePoint = Point_res5.Value;
                        Point_res6 = Editor1.GetPoint(PP6);

                        if (Point_res6.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;

                            return;
                        }

                        Point3d pt5_picked = new Point3d(Point_res5.Value.X, Point_res5.Value.Y, Slope_poly.Elevation);
                        Point3d pt6_picked = new Point3d(Point_res6.Value.X, Point_res6.Value.Y, Slope_poly.Elevation);

                        Point3d pt5 = Slope_poly.GetClosestPointTo(pt5_picked, Vector3d.ZAxis, false);
                        Point3d pt6 = Slope_poly.GetClosestPointTo(pt6_picked, Vector3d.ZAxis, false);

                        double Sta_slope5 = Slope_poly.GetDistAtPoint(pt5);
                        double Sta_slope6 = Slope_poly.GetDistAtPoint(pt6);

                        if (Sta_slope6 < Sta_slope5)
                        {
                            Point3d t = pt5;
                            pt5 = pt6;
                            pt6 = t;
                            double tt = Sta_slope5;
                            Sta_slope5 = Sta_slope6;
                            Sta_slope6 = tt;
                        }

                        double param5 = Slope_poly.GetParameterAtDistance(Sta_slope5);
                        double param6 = Slope_poly.GetParameterAtDistance(Sta_slope6);

                        int index1 = 0;
                        Polyline partial_poly = new Polyline();
                        partial_poly.AddVertexAt(index1, new Point2d(pt5.X, pt5.Y), 0, 0, 0);
                        ++index1;
                        if (Math.Ceiling(param5) < Math.Floor(param6))
                        {
                            for (int k = Convert.ToInt32(Math.Ceiling(param5)); k <= Math.Floor(param6); ++k)
                            {
                                partial_poly.AddVertexAt(index1, Slope_poly.GetPoint2dAt(k), 0, 0, 0);
                                ++index1;
                            }
                        }
                        partial_poly.AddVertexAt(index1, new Point2d(pt6.X, pt6.Y), 0, 0, 0);
                        partial_poly.Elevation = Prof_poly.Elevation;

                        int slope_index = Convert.ToInt32(Math.Ceiling(param5));

                        for (int i = 0; i < partial_poly.NumberOfVertices - 1; ++i)
                        {
                            Point3d pt1 = partial_poly.GetPointAtParameter(i);
                            Point3d pt2 = partial_poly.GetPointAtParameter(i + 1);

                            double sta01 = sta0 + pt1.X - Prof_poly.StartPoint.X;
                            double sta02 = sta0 + pt2.X - Prof_poly.StartPoint.X;
                            double slope = calc_slope(pt1, pt2);

                            if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                            {
                                double min_dist = 1000000;
                                if (Functions.IsNumeric(textBox_min_berm.Text) == true)
                                {
                                    min_dist = Convert.ToDouble(textBox_min_berm.Text);
                                }

                                double cumul_dist = 0;
                                for (int k = 0; k < dt_terrain.Rows.Count; ++k)
                                {
                                    if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value)
                                    {
                                        double t1 = Convert.ToDouble(dt_terrain.Rows[k][0]);
                                        double t2 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                        int category = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                        if (sta01 >= t1 && sta02 <= t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = pt1;
                                            xline1.SecondPoint = new Point3d(pt1.X, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = pt2;
                                            xline2.SecondPoint = new Point3d(pt2.X, pt2.Y + 10, pt2.Z);
                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;
                                                cumul_dist = dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + slope_index, category, pt1, ln);
                                            }
                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;
                                        }
                                        else if (sta01 > t1 && sta02 > t2 && sta01 < t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = pt1;
                                            xline1.SecondPoint = new Point3d(pt1.X, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = new Point3d(pt1.X + t2 - sta01, pt1.Y, pt1.Z);
                                            xline2.SecondPoint = new Point3d(pt1.X + t2 - sta01, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + slope_index, category, pt1, ln);
                                            }

                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;


                                        }
                                        else if (sta01 < t1 && sta02 < t2 && sta02 > t1)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = new Point3d(pt1.X + t1 - sta01, pt1.Y, pt1.Z);
                                            xline1.SecondPoint = new Point3d(pt1.X + t1 - sta01, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = pt2;
                                            xline2.SecondPoint = new Point3d(pt2.X, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + slope_index, category, pt1, ln);
                                            }

                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;

                                        }
                                        else if (sta01 <= t1 && sta02 >= t2)
                                        {
                                            Xline xline1 = new Xline();
                                            Xline xline2 = new Xline();
                                            xline1.BasePoint = new Point3d(pt1.X + t1 - sta01, pt1.Y, pt1.Z);
                                            xline1.SecondPoint = new Point3d(pt1.X + t1 - sta01, pt1.Y + 10, pt1.Z);
                                            xline2.BasePoint = new Point3d(pt1.X + t2 - sta01, pt1.Y, pt1.Z);
                                            xline2.SecondPoint = new Point3d(pt1.X + t2 - sta01, pt2.Y + 10, pt2.Z);

                                            Point3dCollection colint1 = Functions.Intersect_on_both_operands(Prof_poly, xline1);
                                            Point3dCollection colint2 = Functions.Intersect_on_both_operands(Prof_poly, xline2);
                                            if (colint1.Count > 0 && colint2.Count > 0)
                                            {
                                                double sta1 = sta0 + colint1[0].X - Prof_poly.StartPoint.X;
                                                double sta2 = sta0 + colint2[0].X - Prof_poly.StartPoint.X;
                                                double dist = sta2 - sta1;

                                                cumul_dist = cumul_dist + dist;

                                                insert_div_berm_block(ThisDrawing, BTrecord, colint1[0], colint2[0], sta0, cumul_dist, min_dist, slope, i + slope_index, category, pt1, ln);
                                            }



                                        }
                                        else if (sta02 < t1 && sta02 < t2)
                                        {
                                            cumul_dist = 0;
                                            k = dt_terrain.Rows.Count;
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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }

        private void Button_recalculate_position_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;




            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            if (dt_terrain == null || dt_terrain.Rows.Count == 0)
            {
                MessageBox.Show("No terrains loaded. Load the terrain table first");
                return;
            }

            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                set_enable_false();
                this.WindowState = FormWindowState.Minimized;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {



                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect diversion berms:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Editor1.SetImpliedSelection(Empty_array);

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_poly1.MessageForAdding = "\nselect optimized slope polyline:";
                        Prompt_poly1.SingleOnly = true;


                        Rezultat_poly1 = null;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                        Prompt_poly2.SingleOnly = true;


                        Rezultat_poly2 = null;
                        if (slope_id == ObjectId.Null || profile_id == ObjectId.Null)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Rezultat_poly1 = ThisDrawing.Editor.GetSelection(Prompt_poly1);
                            if (Rezultat_poly1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                set_enable_true();
                                panel_poly_profile.BackColor = Color.Red;
                                panel_poly_slope.BackColor = Color.Red;
                                slope_id = ObjectId.Null;
                                profile_id = ObjectId.Null;
                                this.WindowState = FormWindowState.Normal;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            Editor1.SetImpliedSelection(Empty_array);
                            Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);
                            if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                set_enable_true();
                                panel_poly_profile.BackColor = Color.Red;
                                panel_poly_slope.BackColor = Color.Red;
                                slope_id = ObjectId.Null;
                                profile_id = ObjectId.Null;
                                this.WindowState = FormWindowState.Normal;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }


                        }

                        if (slope_id == ObjectId.Null)
                        {
                            slope_id = Rezultat_poly1.Value[0].ObjectId;
                        }

                        if (profile_id == ObjectId.Null)
                        {
                            profile_id = Rezultat_poly2.Value[0].ObjectId;
                        }

                        Entity Ent1 = Trans1.GetObject(slope_id, OpenMode.ForRead) as Entity;

                        if ((Ent1 is Polyline) == false)
                        {
                            MessageBox.Show("the optimized slope is not a polyline\r\n" + Ent1.GetType().ToString() + "\r\nOperation aborted");
                            this.WindowState = FormWindowState.Normal;
                            set_enable_true();

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Entity Ent2 = Trans1.GetObject(profile_id, OpenMode.ForRead) as Entity;
                        if ((Ent2 is Polyline) == false)
                        {
                            set_enable_true();
                            this.WindowState = FormWindowState.Normal;
                            MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent2.GetType().ToString() + "\r\nOperation aborted");
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Slope_poly = Ent1 as Polyline;
                        Prof_poly = Ent2 as Polyline;

                        string note1 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + "Hr" + DateTime.Now.Minute.ToString() + "Min_by" + Environment.UserName.ToUpper();

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Entity Ent3 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForWrite) as Entity;
                            if (Ent3 is BlockReference)
                            {
                                BlockReference block1 = Ent3 as BlockReference;
                                if (block1.AttributeCollection.Count > 0)
                                {

                                    Xline xline3 = new Xline();
                                    xline3.BasePoint = new Point3d(block1.Position.X, block1.Position.Y, Slope_poly.Elevation);
                                    xline3.SecondPoint = new Point3d(block1.Position.X, block1.Position.Y + 10, Slope_poly.Elevation);

                                    Point3dCollection col3 = Functions.Intersect_on_both_operands(xline3, Slope_poly);
                                    Point3dCollection colint3 = Functions.Intersect_on_both_operands(xline3, Prof_poly);

                                    string slopeid = "XX";
                                    if (col3.Count > 0)
                                    {
                                        double param3 = Slope_poly.GetParameterAtPoint(col3[0]);

                                        slopeid = Convert.ToString(Math.Ceiling(param3));
                                    }
                                    xline3.Dispose();

                                    double X = block1.Position.X;
                                    double Y = block1.Position.Y;
                                    double Z = Prof_poly.Elevation;
                                    double staX = sta0 + X - Prof_poly.StartPoint.X;


                                    if (dt_prof != null && dt_prof.Rows.Count > 0 && colint3.Count > 0)
                                    {
                                        double param1 = Prof_poly.GetParameterAtPoint(Prof_poly.GetClosestPointTo(colint3[0], Vector3d.ZAxis, false));
                                        int idx0 = Convert.ToInt32(Math.Floor(param1));
                                        double dif1 = Prof_poly.GetPointAtParameter(param1).X - Prof_poly.GetPointAtParameter(idx0).X;

                                        if (dt_prof.Rows.Count >= idx0 + 1)
                                        {
                                            if (dt_prof.Rows[idx0][Col_sta] != DBNull.Value)
                                            {
                                                double sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta]);
                                                if (dt_prof.Rows[idx0][Col_sta_eq] != DBNull.Value)
                                                {
                                                    sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta_eq]);
                                                }
                                                staX = sta_start + dif1;
                                            }
                                        }
                                    }

                                    string sta1 = Functions.Get_chainage_from_double(staX, "m", 1);

                                    string cat1 = "terrain not loaded";

                                    if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                                    {
                                        for (int k = 0; k < dt_terrain.Rows.Count; ++k)
                                        {
                                            if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value)
                                            {
                                                double t1 = Convert.ToDouble(dt_terrain.Rows[k][0]);
                                                double t2 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                                int category = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                                string cat2 = "terrain not defined";

                                                if (category == 1)
                                                {
                                                    cat2 = "Fine Sand";
                                                }
                                                if (category == 2)
                                                {
                                                    cat2 = "Clay";
                                                }
                                                if (category == 3)
                                                {
                                                    cat2 = "Gravel/Bedrock";
                                                }

                                                if (staX >= t1 && staX <= t2)
                                                {
                                                    cat1 = cat2;
                                                    k = dt_terrain.Rows.Count;
                                                }

                                            }
                                        }
                                    }

                                    for (int j = 0; j < block1.AttributeCollection.Count; ++j)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(block1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                        if (atr1 != null)
                                        {
                                            if (atr1.Tag.ToLower() == "sta")
                                            {
                                                atr1.TextString = sta1;
                                            }
                                            if (atr1.Tag.ToLower() == "terrain")
                                            {
                                                atr1.TextString = cat1;
                                            }
                                            if (atr1.Tag.ToLower() == "slopeid")
                                            {
                                                atr1.TextString = slopeid;
                                            }
                                            if (atr1.Tag.ToLower() == "notes")
                                            {
                                                atr1.TextString = note1;
                                            }
                                        }
                                    }

                                    Xline xline1 = new Xline();
                                    xline1.BasePoint = new Point3d(X, Y, Z);
                                    xline1.SecondPoint = new Point3d(X, Y + 10, Z);
                                    Point3dCollection colint = Functions.Intersect_on_both_operands(Prof_poly, xline1);

                                    if (colint.Count > 0)
                                    {
                                        double y1 = colint[0].Y;
                                        block1.TransformBy(Matrix3d.Displacement(block1.Position.GetVectorTo(new Point3d(X, y1, Z))));

                                    }
                                }
                            }
                        }

                        Update_blocks_slopeID(Trans1, BTrecord, Slope_poly);
                        Update_one_of_one(Trans1, BTrecord, Slope_poly);
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
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }

        private void Update_blocks_slopeID(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, BlockTableRecord BTrecord, Polyline poly_slope)
        {
            if (Trans1 != null && BTrecord != null && poly_slope != null)
            {
                dt_blocks = new System.Data.DataTable();
                dt_blocks.Columns.Add("objectid", typeof(ObjectId));
                dt_blocks.Columns.Add("sta", typeof(double));
                dt_blocks.Columns.Add("slopeid", typeof(int));
                dt_blocks.Columns.Add("Slope", typeof(double));


                foreach (ObjectId odid1 in BTrecord)
                {
                    BlockReference block1 = Trans1.GetObject(odid1, OpenMode.ForWrite) as BlockReference;
                    if (block1 != null && block1.AttributeCollection.Count > 0)
                    {
                        string name1 = Functions.get_block_name(block1);
                        if (block1.Layer == "DivBerm" && name1 == "DB")
                        {
                            Xline xline3 = new Xline();
                            xline3.BasePoint = new Point3d(block1.Position.X, block1.Position.Y, poly_slope.Elevation);
                            xline3.SecondPoint = new Point3d(block1.Position.X, block1.Position.Y + 10, poly_slope.Elevation);
                            Point3dCollection colint3 = Functions.Intersect_on_both_operands(poly_slope, xline3);
                            if (colint3.Count > 0)
                            {
                                double param3 = poly_slope.GetParameterAtPoint(colint3[0]);

                                LineSegment3d line1 = poly_slope.GetLineSegmentAt(Convert.ToInt32(Math.Floor(param3)));

                                double Slope1 = calc_slope(line1.StartPoint, line1.EndPoint);

                                dt_blocks.Rows.Add();
                                dt_blocks.Rows[dt_blocks.Rows.Count - 1][0] = block1.ObjectId;
                                dt_blocks.Rows[dt_blocks.Rows.Count - 1][1] = block1.Position.X;
                                dt_blocks.Rows[dt_blocks.Rows.Count - 1][2] = Convert.ToInt32(Math.Ceiling(param3));
                                dt_blocks.Rows[dt_blocks.Rows.Count - 1][3] = Math.Round(Slope1, 1);

                                for (int j = 0; j < block1.AttributeCollection.Count; ++j)
                                {
                                    AttributeReference atr1 = Trans1.GetObject(block1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                    if (atr1 != null)
                                    {
                                        if (atr1.Tag.ToLower() == "slopeid")
                                        {
                                            atr1.TextString = Convert.ToString(Math.Ceiling(param3));
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
                dt_blocks = Functions.Sort_data_table(dt_blocks, "sta");
            }
        }



        private void Update_one_of_one(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, BlockTableRecord BTrecord, Polyline poly_slope)
        {
            if (Trans1 != null && BTrecord != null && poly_slope != null)
            {
                if (dt_blocks != null && dt_blocks.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_blocks.Rows.Count; ++i)
                    {
                        if (dt_blocks.Rows[i][0] != DBNull.Value && dt_blocks.Rows[i][1] != DBNull.Value && dt_blocks.Rows[i][2] != DBNull.Value)
                        {
                            try
                            {
                                ObjectId objectid1 = (ObjectId)dt_blocks.Rows[i][0];
                                int slopeid1 = Convert.ToInt32(dt_blocks.Rows[i][2]);
                                double slope1 = Convert.ToDouble(dt_blocks.Rows[i][3]);
                                double sta = Convert.ToDouble(dt_blocks.Rows[i]["sta"]);

                                int terrain = 1;

                                if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                                {
                                    for (int j = 0; j < dt_terrain.Rows.Count; ++j)
                                    {
                                        double sta1 = Convert.ToDouble(dt_terrain.Rows[j]["sta1"]);
                                        double sta2 = Convert.ToDouble(dt_terrain.Rows[j]["sta2"]);
                                        if (sta1 <= sta && sta >= sta2)
                                        {
                                            terrain = Convert.ToInt32(dt_terrain.Rows[j]["soil_category"]);
                                            j = dt_terrain.Rows.Count;
                                        }
                                    }
                                }

                                BlockReference block1 = Trans1.GetObject(objectid1, OpenMode.ForRead) as BlockReference;

                                List<ObjectId> lista1 = new List<ObjectId>();
                                List<int> lista2 = new List<int>();
                                List<int> lista3 = new List<int>();
                               

                                int index1 = 1;
                                int total = 1;




                                lista1.Add(objectid1);
                                lista2.Add(index1);
                                lista3.Add(total);
                              



                                if (block1 != null && i < dt_blocks.Rows.Count - 1)
                                {
                                    for (int j = i + 1; j < dt_blocks.Rows.Count; ++j)
                                    {
                                        if (dt_blocks.Rows[j][0] != DBNull.Value && dt_blocks.Rows[j][1] != DBNull.Value && dt_blocks.Rows[j][2] != DBNull.Value)
                                        {
                                            ObjectId objectid2 = (ObjectId)dt_blocks.Rows[j][0];
                                            int slopeid2 = Convert.ToInt32(dt_blocks.Rows[j][2]);
                                            if (slopeid1 == slopeid2)
                                            {
                                                BlockReference block2 = Trans1.GetObject(objectid2, OpenMode.ForRead) as BlockReference;
                                                ++index1;
                                                ++total;
                                                lista1.Add(objectid2);
                                                lista2.Add(index1);

                                                for (int k = 0; k < lista3.Count; ++k)
                                                {
                                                    lista3[k] = total;
                                                }

                                                lista3.Add(total);

                                            }
                                            else
                                            {
                                                j = dt_blocks.Rows.Count;
                                            }
                                        }
                                    }
                                  


                                    for (int k = 0; k < lista3.Count; ++k)
                                    {
                                        BlockReference block3 = Trans1.GetObject(lista1[k], OpenMode.ForWrite) as BlockReference;
                                        if (block3 != null && block3.AttributeCollection.Count > 0)
                                        {

                                            for (int n = 0; n < block3.AttributeCollection.Count; ++n)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(block3.AttributeCollection[n], OpenMode.ForWrite) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    if (atr1.Tag.ToLower() == "no")
                                                    {
                                                        atr1.TextString = Convert.ToString(lista2[k]) + " of " + Convert.ToString(lista3[k]);
                                                    }

                                                    if (atr1.Tag.ToLower() == "ditchplug")
                                                    {
                                                        atr1.TextString = "NO";
                                                    }


                                                    if (atr1.Tag.ToLower() == "slope")
                                                    {
                                                        atr1.TextString = "AVG SLOPE = " + Functions.Get_String_Rounded(slope1, 1) + "%";
                                                    }

                                                }
                                            }
                                        }
                                    }


                                }


                                if (block1 != null && i == dt_blocks.Rows.Count - 1)
                                {
                                    for (int k = 0; k < lista3.Count; ++k)
                                    {
                                        BlockReference block3 = Trans1.GetObject(lista1[k], OpenMode.ForWrite) as BlockReference;
                                        if (block3 != null && block3.AttributeCollection.Count > 0)
                                        {

                                            for (int n = 0; n < block3.AttributeCollection.Count; ++n)
                                            {
                                                AttributeReference atr1 = Trans1.GetObject(block3.AttributeCollection[n], OpenMode.ForWrite) as AttributeReference;
                                                if (atr1 != null)
                                                {
                                                    if (atr1.Tag.ToLower() == "no")
                                                    {
                                                        atr1.TextString = Convert.ToString(lista2[k]) + " of " + Convert.ToString(lista3[k]);
                                                    }

                                                    if (atr1.Tag.ToLower() == "ditchplug")
                                                    {
                                                        atr1.TextString = "NO";
                                                    }


                                                    if (atr1.Tag.ToLower() == "slope")
                                                    {
                                                        atr1.TextString = "AVG SLOPE = " + Functions.Get_String_Rounded(slope1, 1) + "%";
                                                    }

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
                    }
                }
            }

        }

        private void button_export_data_to_excel_Click(object sender, EventArgs e)
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

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect diversion berms:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }

                        if (Rezultat1.Status == PromptStatus.OK)
                        {
                            System.Data.DataTable dt1 = new System.Data.DataTable();

                            dt1.Columns.Add("BlockName", typeof(string));
                            dt1.Columns.Add("Layer", typeof(string));
                            dt1.Columns.Add("Station", typeof(string));
                            dt1.Columns.Add("north", typeof(string));
                            dt1.Columns.Add("east", typeof(string));
                            dt1.Columns.Add("Spacing", typeof(string));
                            dt1.Columns.Add("DitchPlug", typeof(string));
                            dt1.Columns.Add("Slope", typeof(string));
                            dt1.Columns.Add("SlopeItem", typeof(string));
                            dt1.Columns.Add("Terrain", typeof(string));
                            dt1.Columns.Add("Extend existing BERM", typeof(string));




                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;

                                if (block1 != null && block1.AttributeCollection.Count > 0)
                                {
                                    string block_name = Functions.get_block_name(block1);
                                    if (block_name.ToUpper() == "DB")
                                    {
                                        string sta = "";
                                        string slope = "";
                                        string terrain = "";
                                        string one_of_one = "";

                                        string layer = block1.Layer;
                                        string yesno = "NO";
                                        string spacing = "";




                                        for (int j = 0; j < block1.AttributeCollection.Count; ++j)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(block1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                if (atr1.Tag.ToLower() == "sta")
                                                {
                                                    sta = atr1.TextString;
                                                }
                                                if (atr1.Tag.ToLower() == "terrain")
                                                {
                                                    terrain = atr1.TextString;
                                                }
                                                if (atr1.Tag.ToLower() == "slope")
                                                {
                                                    slope = atr1.TextString;
                                                }
                                                if (atr1.Tag.ToLower() == "no")
                                                {
                                                    one_of_one = atr1.TextString;

                                                }
                                                if (atr1.Tag.ToUpper() == "DITCHPLUG")
                                                {
                                                    yesno = atr1.TextString;

                                                    if (yesno.ToLower() == "yes" || yesno.ToLower() == "true" || yesno.ToLower() == "y")
                                                    {
                                                        yesno = "Y";
                                                    }
                                                    else
                                                    {
                                                        yesno = "N";
                                                    }

                                                }
                                                if (atr1.Tag.ToUpper() == "SPACING")
                                                {
                                                    spacing = atr1.TextString;
                                                }
                                            }
                                        }
                                        if (slope != "")
                                        {
                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1]["BlockName"] = block_name;
                                            dt1.Rows[dt1.Rows.Count - 1]["Layer"] = layer;
                                            dt1.Rows[dt1.Rows.Count - 1]["Slope"] = slope;
                                            dt1.Rows[dt1.Rows.Count - 1]["SlopeItem"] = one_of_one;
                                            dt1.Rows[dt1.Rows.Count - 1]["Terrain"] = terrain;
                                            dt1.Rows[dt1.Rows.Count - 1]["Station"] = sta;
                                            dt1.Rows[dt1.Rows.Count - 1]["Spacing"] = spacing;
                                            dt1.Rows[dt1.Rows.Count - 1]["DitchPlug"] = yesno;
                                        }
                                    }

                                }
                            }


                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;

                                if (block1 != null )
                                {
                                    string block_name = Functions.get_block_name(block1);
                                    if (block_name.ToUpper() == "DB-EX")
                                    {
                                        string sta = "";





                                        for (int j = 0; j < block1.AttributeCollection.Count; ++j)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(block1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                if (atr1.Tag.ToLower() == "sta")
                                                {
                                                    sta = atr1.TextString;

                                                    j = block1.AttributeCollection.Count;

                                                }
                                               
                                            }
                                        }
                                        if (sta != "")
                                        {
                                            for (int l = 0; l < dt1.Rows.Count; ++l)
                                            {
                                                if (Convert.ToString(dt1.Rows[l]["Station"]).ToLower().Replace(" ","")==sta.ToLower())
                                                {
                                                    dt1.Rows[l]["Extend existing BERM"] = "YES";
                                                }
                                            }

                                        }
                                    }

                                }
                            }


                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
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
            this.WindowState = FormWindowState.Normal;

        }

        private void button_load_profile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    string File1 = fbd.FileName;
                    Load_existing_profile_graph(File1);
                    if (dt_prof != null && dt_prof.Rows.Count > 0)
                    {
                        double start_sta = 0;
                        if (dt_prof.Rows[0][Col_sta] != DBNull.Value)
                        {
                            start_sta = Convert.ToDouble(dt_prof.Rows[0][Col_sta]);
                            if (dt_prof.Rows[0][Col_sta_eq] != DBNull.Value)
                            {
                                start_sta = Convert.ToDouble(dt_prof.Rows[0][Col_sta_eq]);
                            }

                        }

                        textBox_start_prof_sta.Text = Convert.ToString(start_sta);
                        label_prof_loaded.Text = System.IO.Path.GetFileName(File1);
                        label_prof_loaded.ForeColor = Color.Green;
                    }

                }
            }

            if (dt_prof == null || dt_prof.Rows.Count == 0)
            {
                label_prof_loaded.Text = "Profile.xlsx not loaded";
                label_prof_loaded.ForeColor = Color.Red;
                dt_prof = null;
            }
        }

        public System.Data.DataTable Load_existing_profile_graph(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the profile data file does not exist");
                return null;
            }


            dt_prof = new System.Data.DataTable();

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt_prof = Build_Data_table_profile_from_excel(W1);

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
            return dt_prof;

        }

        public System.Data.DataTable Build_Data_table_profile_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1)
        {


            System.Data.DataTable Data_table_profile = Creaza_profile_datatable_structure();


            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["D9:D100008"];
            object[,] values2 = new object[100000, 1];
            values2 = range2.Value2;

            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    Data_table_profile.Rows.Add();

                }
                else
                {
                    i = values2.Length + 1;
                }
            }

            int NrR = Data_table_profile.Rows.Count;
            int NrC = Data_table_profile.Columns.Count;




            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A9:G" + Convert.ToString(9 + NrR - 1)];





            object[,] values = new object[NrR - 1, NrC - 1];

            values = range1.Value2;

            for (int i = 0; i < Data_table_profile.Rows.Count; ++i)
            {
                for (int j = 0; j < Data_table_profile.Columns.Count; ++j)
                {
                    object Valoare = values[i + 1, j + 1];
                    if (Valoare == null) Valoare = DBNull.Value;

                    Data_table_profile.Rows[i][j] = Valoare;
                }
            }




            return Data_table_profile;


        }



        public System.Data.DataTable Creaza_profile_datatable_structure()
        {

            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_sta);
            Lista1.Add(Col_sta_eq);
            Lista1.Add(Col_Elev);
            Lista1.Add(Col_Type);
            Lista1.Add(Col_Elev1);
            Lista1.Add(Col_Elev2);

            Lista2.Add(typeof(string));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(string));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(double));

            System.Data.DataTable Data_table_prof = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                Data_table_prof.Columns.Add(Lista1[i], Lista2[i]);
            }
            return Data_table_prof;
        }

        private void button_place_berm_based_on_point_Click(object sender, EventArgs e)
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

                        string blockname = "DB";
                        double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);
                        if (blockname != "")
                        {
                            if (BlockTable1.Has(blockname) == true)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                this.WindowState = FormWindowState.Minimized;

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_poly1.MessageForAdding = "\nselect optimized slope polyline:";
                                Prompt_poly1.SingleOnly = true;
                                Rezultat_poly1 = ThisDrawing.Editor.GetSelection(Prompt_poly1);

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                                Prompt_poly2.SingleOnly = true;

                                Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                                if (Rezultat_poly1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    set_enable_true();

                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    set_enable_true();
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }


                                Entity Ent1 = Trans1.GetObject(Rezultat_poly1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                if ((Ent1 is Polyline) == false)
                                {
                                    MessageBox.Show("the optimized slope is not a polyline\r\n" + Ent1.GetType().ToString() + "\r\nOperation aborted");
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                Entity Ent2 = Trans1.GetObject(Rezultat_poly2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                if ((Ent2 is Polyline) == false)
                                {
                                    set_enable_true();
                                    this.WindowState = FormWindowState.Normal;

                                    MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent2.GetType().ToString() + "\r\nOperation aborted");
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                Polyline Slope_poly = Ent1 as Polyline;
                                Polyline Prof_poly = Ent2 as Polyline;


                                BlockTableRecord btr = Trans1.GetObject(BlockTable1[blockname], OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null)
                                {
                                    string layer1 = "0";
                                    double scale1 = 1;

                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                        if (block1 != null)
                                        {
                                            if (Functions.get_block_name(block1) == blockname)
                                            {
                                                layer1 = block1.Layer;
                                                scale1 = block1.ScaleFactors.X;
                                                break;
                                            }
                                        }
                                    }


                                    BlockReference br1 = null;
                                    jig_actions.insert_block(ref br1, blockname, layer1, Prof_poly, scale1);




                                    string note1 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + "Hr" + DateTime.Now.Minute.ToString() + "Min_by" + Environment.UserName.ToUpper();


                                    Xline xline3 = new Xline();
                                    xline3.BasePoint = new Point3d(br1.Position.X, br1.Position.Y, Slope_poly.Elevation);
                                    xline3.SecondPoint = new Point3d(br1.Position.X, br1.Position.Y + 10, Slope_poly.Elevation);

                                    Point3dCollection col3 = Functions.Intersect_on_both_operands(xline3, Slope_poly);
                                    Point3dCollection colint3 = Functions.Intersect_on_both_operands(xline3, Prof_poly);

                                    string slopeid = "XX";
                                    double slope = 0;
                                    if (col3.Count > 0)
                                    {
                                        double param3 = Slope_poly.GetParameterAtPoint(col3[0]);

                                        double index1 = Math.Floor(param3);

                                        Point3d pt1 = Slope_poly.GetPointAtParameter(index1);
                                        Point3d pt2 = Slope_poly.GetPointAtParameter(index1 + 1);

                                        slope = calc_slope(pt1, pt2);


                                        slopeid = Convert.ToString(Math.Ceiling(param3));
                                    }
                                    xline3.Dispose();



                                    double sta1 = sta0 + br1.Position.X - Prof_poly.StartPoint.X;


                                    if (dt_prof != null && dt_prof.Rows.Count > 0 && colint3.Count > 0)
                                    {
                                        double param1 = Prof_poly.GetParameterAtPoint(Prof_poly.GetClosestPointTo(colint3[0], Vector3d.ZAxis, false));
                                        int idx0 = Convert.ToInt32(Math.Floor(param1));
                                        double dif1 = Prof_poly.GetPointAtParameter(param1).X - Prof_poly.GetPointAtParameter(idx0).X;

                                        if (dt_prof.Rows.Count >= idx0 + 1)
                                        {
                                            if (dt_prof.Rows[idx0][Col_sta] != DBNull.Value)
                                            {
                                                double sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta]);
                                                if (dt_prof.Rows[idx0][Col_sta_eq] != DBNull.Value)
                                                {
                                                    sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta_eq]);
                                                }
                                                sta1 = sta_start + dif1;
                                            }
                                        }
                                    }

                                    string sta_string = Functions.Get_chainage_from_double(sta1, "m", 1);


                                    string slope_string = "AVG SLOPE = " + Functions.Get_String_Rounded(slope, 1) + "%";

                                    string cat1 = "terrain not loaded";
                                    double spacing = 0;
                                    if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                                    {
                                        for (int k = 0; k < dt_terrain.Rows.Count; ++k)
                                        {
                                            if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value)
                                            {
                                                double t1 = Convert.ToDouble(dt_terrain.Rows[k][0]);
                                                double t2 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                                int category = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                                string cat2 = "terrain not defined";

                                                if (category == 1)
                                                {
                                                    cat2 = "Fine Sand";
                                                }
                                                if (category == 2)
                                                {
                                                    cat2 = "Clay";
                                                }
                                                if (category == 3)
                                                {
                                                    cat2 = "Gravel/Bedrock";
                                                }

                                                if (sta1 >= t1 && sta1 <= t2)
                                                {
                                                    cat1 = cat2;
                                                    k = dt_terrain.Rows.Count;
                                                    spacing = get_spacing(Math.Abs(slope), category);
                                                }

                                            }
                                        }





                                    }

                                    string spacing_string = "SPACING = " + Functions.Get_String_Rounded(spacing, 1);

                                    for (int j = 0; j < br1.AttributeCollection.Count; ++j)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(br1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                        if (atr1 != null)
                                        {
                                            if (atr1.Tag.ToLower() == "sta")
                                            {
                                                atr1.TextString = sta_string;
                                            }
                                            if (atr1.Tag.ToLower() == "terrain")
                                            {
                                                atr1.TextString = cat1;
                                            }
                                            if (atr1.Tag.ToLower() == "slopeid")
                                            {
                                                atr1.TextString = slopeid;
                                            }
                                            if (atr1.Tag.ToLower() == "notes")
                                            {
                                                atr1.TextString = note1;
                                            }
                                            if (atr1.Tag.ToLower() == "slope")
                                            {
                                                atr1.TextString = slope_string;
                                            }
                                            if (atr1.Tag.ToLower() == "spacing")
                                            {
                                                atr1.TextString = spacing_string;
                                            }
                                        }
                                    }

                                    Update_blocks_slopeID(Trans1, BTrecord, Slope_poly);
                                    Update_one_of_one(Trans1, BTrecord, Slope_poly);

                                    Trans1.Commit();
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            this.WindowState = FormWindowState.Normal;
        }

        private void button_place_berm_based_on_sta_Click(object sender, EventArgs e)
        {

            string text_string = textBox_div_berm_sta.Text;
            if (Functions.IsNumeric(text_string.Replace("+", "")) == false) return;

            double sta1 = Convert.ToDouble(text_string.Replace("+", ""));

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

                        string blockname = "DB";
                        double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);
                        if (blockname != "")
                        {
                            if (BlockTable1.Has(blockname) == true)
                            {
                                Editor1.SetImpliedSelection(Empty_array);
                                this.WindowState = FormWindowState.Minimized;

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly1 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_poly1.MessageForAdding = "\nselect optimized slope polyline:";
                                Prompt_poly1.SingleOnly = true;
                                Rezultat_poly1 = ThisDrawing.Editor.GetSelection(Prompt_poly1);

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                                Prompt_poly2.SingleOnly = true;

                                Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                                if (Rezultat_poly1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    set_enable_true();

                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    set_enable_true();
                                    this.WindowState = FormWindowState.Normal;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }


                                Entity Ent1 = Trans1.GetObject(Rezultat_poly1.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                if ((Ent1 is Polyline) == false)
                                {
                                    MessageBox.Show("the optimized slope is not a polyline\r\n" + Ent1.GetType().ToString() + "\r\nOperation aborted");
                                    this.WindowState = FormWindowState.Normal;
                                    set_enable_true();

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                Entity Ent2 = Trans1.GetObject(Rezultat_poly2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                if ((Ent2 is Polyline) == false)
                                {
                                    set_enable_true();
                                    this.WindowState = FormWindowState.Normal;

                                    MessageBox.Show("the polyline profile is not a polyline\r\n" + Ent2.GetType().ToString() + "\r\nOperation aborted");
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                Polyline Slope_poly = Ent1 as Polyline;
                                Polyline Prof_poly = Ent2 as Polyline;


                                BlockTableRecord btr = Trans1.GetObject(BlockTable1[blockname], OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null)
                                {
                                    string layer1 = "0";
                                    double scale1 = 1;

                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                        if (block1 != null)
                                        {
                                            if (Functions.get_block_name(block1) == blockname)
                                            {
                                                layer1 = block1.Layer;
                                                scale1 = block1.ScaleFactors.X;
                                                break;
                                            }
                                        }
                                    }





                                    double x = Prof_poly.StartPoint.X + sta0 + sta1;

                                    string note1 = DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString() + "/" + DateTime.Now.Year.ToString() + "_" + DateTime.Now.Hour.ToString() + "Hr" + DateTime.Now.Minute.ToString() + "Min_by" + Environment.UserName.ToUpper();


                                    Xline xline3 = new Xline();
                                    xline3.BasePoint = new Point3d(x, 0, Slope_poly.Elevation);
                                    xline3.SecondPoint = new Point3d(x, 10, Slope_poly.Elevation);

                                    Point3dCollection col3 = Functions.Intersect_on_both_operands(xline3, Slope_poly);
                                    Point3dCollection colint3 = Functions.Intersect_on_both_operands(xline3, Prof_poly);

                                    string slopeid = "XX";
                                    double slope = 0;
                                    if (col3.Count > 0)
                                    {
                                        double param3 = Slope_poly.GetParameterAtPoint(col3[0]);

                                        double index1 = Math.Floor(param3);

                                        Point3d pt1 = Slope_poly.GetPointAtParameter(index1);
                                        Point3d pt2 = Slope_poly.GetPointAtParameter(index1 + 1);

                                        slope = calc_slope(pt1, pt2);



                                        slopeid = Convert.ToString(Math.Ceiling(param3));
                                    }
                                    xline3.Dispose();

                                    if (colint3.Count == 0)
                                    {
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        this.WindowState = FormWindowState.Normal;
                                        return;
                                    }

                                    Point3d pt_ins = colint3[0];

                                    BlockReference br1 = new BlockReference(pt_ins, btr.ObjectId);
                                    br1.Layer = layer1;
                                    br1.ScaleFactors = new Scale3d(scale1, scale1, scale1);
                                    br1.ColorIndex = 256;
                                    BTrecord.AppendEntity(br1);
                                    Trans1.AddNewlyCreatedDBObject(br1, true);

                                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = br1.AttributeCollection;
                                    BlockTableRecordEnumerator btr_enum = btr.GetEnumerator();
                                    while (btr_enum.MoveNext())
                                    {
                                        Entity att1 = (Entity)Trans1.GetObject(btr_enum.Current, OpenMode.ForWrite);
                                        if (att1 is AttributeDefinition)
                                        {
                                            AttributeDefinition Attdef = (AttributeDefinition)att1;
                                            AttributeReference Attref = new AttributeReference();
                                            Attref.SetAttributeFromBlock(Attdef, br1.BlockTransform);

                                            if (Attref != null)
                                            {
                                                attColl.AppendAttribute(Attref);
                                                Trans1.AddNewlyCreatedDBObject(Attref, true);
                                            }
                                        }

                                    }




                                    if (dt_prof != null && dt_prof.Rows.Count > 0 && colint3.Count > 0)
                                    {
                                        double param1 = Prof_poly.GetParameterAtPoint(Prof_poly.GetClosestPointTo(colint3[0], Vector3d.ZAxis, false));
                                        int idx0 = Convert.ToInt32(Math.Floor(param1));
                                        double dif1 = Prof_poly.GetPointAtParameter(param1).X - Prof_poly.GetPointAtParameter(idx0).X;

                                        if (dt_prof.Rows.Count >= idx0 + 1)
                                        {
                                            if (dt_prof.Rows[idx0][Col_sta] != DBNull.Value)
                                            {
                                                double sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta]);
                                                if (dt_prof.Rows[idx0][Col_sta_eq] != DBNull.Value)
                                                {
                                                    sta_start = Convert.ToDouble(dt_prof.Rows[idx0][Col_sta_eq]);
                                                }
                                                sta1 = sta_start + dif1;
                                            }
                                        }
                                    }

                                    string sta_string = Functions.Get_chainage_from_double(sta1, "m", 1);


                                    string slope_string = "AVG SLOPE = " + Functions.Get_String_Rounded(slope, 1) + "%";

                                    string cat1 = "terrain not loaded";
                                    double spacing = 0;
                                    if (dt_terrain != null && dt_terrain.Rows.Count > 0)
                                    {
                                        for (int k = 0; k < dt_terrain.Rows.Count; ++k)
                                        {
                                            if (dt_terrain.Rows[k][0] != DBNull.Value && dt_terrain.Rows[k][1] != DBNull.Value)
                                            {
                                                double t1 = Convert.ToDouble(dt_terrain.Rows[k][0]);
                                                double t2 = Convert.ToDouble(dt_terrain.Rows[k][1]);
                                                int category = Convert.ToInt32(dt_terrain.Rows[k][2]);
                                                string cat2 = "terrain not defined";

                                                if (category == 1)
                                                {
                                                    cat2 = "Fine Sand";
                                                }
                                                if (category == 2)
                                                {
                                                    cat2 = "Clay";
                                                }
                                                if (category == 3)
                                                {
                                                    cat2 = "Gravel/Bedrock";
                                                }

                                                if (sta1 >= t1 && sta1 <= t2)
                                                {
                                                    cat1 = cat2;
                                                    k = dt_terrain.Rows.Count;
                                                    spacing = get_spacing(Math.Abs(slope), category);
                                                }

                                            }
                                        }





                                    }

                                    string spacing_string = "SPACING = " + Functions.Get_String_Rounded(spacing, 1);

                                    for (int j = 0; j < br1.AttributeCollection.Count; ++j)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(br1.AttributeCollection[j], OpenMode.ForWrite) as AttributeReference;
                                        if (atr1 != null)
                                        {
                                            if (atr1.Tag.ToLower() == "sta")
                                            {
                                                atr1.TextString = sta_string;
                                            }
                                            if (atr1.Tag.ToLower() == "terrain")
                                            {
                                                atr1.TextString = cat1;
                                            }
                                            if (atr1.Tag.ToLower() == "slopeid")
                                            {
                                                atr1.TextString = slopeid;
                                            }
                                            if (atr1.Tag.ToLower() == "notes")
                                            {
                                                atr1.TextString = note1;
                                            }
                                            if (atr1.Tag.ToLower() == "slope")
                                            {
                                                atr1.TextString = slope_string;
                                            }
                                            if (atr1.Tag.ToLower() == "spacing")
                                            {
                                                atr1.TextString = spacing_string;
                                            }
                                        }
                                    }

                                    Update_blocks_slopeID(Trans1, BTrecord, Slope_poly);
                                    Update_one_of_one(Trans1, BTrecord, Slope_poly);

                                    Trans1.Commit();
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
            this.WindowState = FormWindowState.Normal;
        }

        private void button_select_slope_Click(object sender, EventArgs e)
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
                        this.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_slope;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_slope;
                        Prompt_slope = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the slope polyline:");
                        Prompt_slope.SetRejectMessage("\nSelect a polyline!");
                        Prompt_slope.AllowNone = true;
                        Prompt_slope.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_slope = ThisDrawing.Editor.GetEntity(Prompt_slope);

                        if (Rezultat_slope.Status != PromptStatus.OK)
                        {
                            panel_poly_slope.BackColor = Color.Red;
                            slope_id = ObjectId.Null;
                            this.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        slope_id = Rezultat_slope.ObjectId;
                        panel_poly_slope.BackColor = Color.Green;

                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.WindowState = FormWindowState.Normal;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        private void button_select_profile_polyline_Click(object sender, EventArgs e)
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
                        this.WindowState = FormWindowState.Minimized;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_profile;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_profile;
                        Prompt_profile = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the profile polyline:");
                        Prompt_profile.SetRejectMessage("\nSelect a polyline!");
                        Prompt_profile.AllowNone = true;
                        Prompt_profile.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_profile = ThisDrawing.Editor.GetEntity(Prompt_profile);

                        if (Rezultat_profile.Status != PromptStatus.OK)
                        {
                            panel_poly_profile.BackColor = Color.Red;
                            profile_id = ObjectId.Null;
                            this.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            Editor1.SetImpliedSelection(Empty_array);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        profile_id = Rezultat_profile.ObjectId;
                        panel_poly_profile.BackColor = Color.Green;

                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.WindowState = FormWindowState.Normal;
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        private void button_place_existing_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }

            if (Functions.IsNumeric(textBox_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }
            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);
            try
            {


                if (comboBox_ws1.Text != "")
                {
                    string string1 = comboBox_ws1.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            set_enable_false();
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null)
                            {
                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("sta", typeof(double));
                                List<string> list_col = new List<string>();
                                list_col.Add("sta");
                                List<string> list_colxl = new List<string>();
                                list_colxl.Add(textBox_existing_station.Text);

                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, list_col, list_colxl);
                                if (dt1.Rows.Count > 0)
                                {

                                    ObjectId[] Empty_array = null;
                                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                                    Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                    {
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                        Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                                        Prompt_poly2.SingleOnly = true;
                                        this.WindowState = FormWindowState.Minimized;
                                        Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                                        if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            this.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {
                                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                            Entity Ent2 = Trans1.GetObject(Rezultat_poly2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                            if ((Ent2 is Polyline) == false)
                                            {
                                                MessageBox.Show("the polyline profile is not a polyline\r\n is a " + Ent2.GetType().ToString() + "\r\nOperation aborted");
                                                this.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }
                                            Polyline poly1 = Ent2 as Polyline;
                                            double ymin = -1000000;
                                            double ymax = 1000000;

                                            for (int i = 0; i < poly1.NumberOfVertices - 1; ++i)
                                            {
                                                double y = poly1.GetPointAtParameter(i).Y;
                                                if (i == 0)
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

                                            string ln1 = "_existing_db";
                                            string block_name = "DB-EX";
                                            Functions.Creaza_layer(ln1, 2, true);

                                          

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                double sta1 = Convert.ToDouble(dt1.Rows[i][0]);



                                                double diferenta1 = 0;
                                           



                                                int index_sta1 = -1;
                                             
                                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][Col_sta] != DBNull.Value && dt_prof.Rows[j + 1][Col_sta] != DBNull.Value)
                                                        {
                                                            double sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta]);
                                                            if (dt_prof.Rows[j][Col_sta_eq] != DBNull.Value)
                                                            {
                                                                sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta_eq]);
                                                            }
                                                            double sta_end = Convert.ToDouble(dt_prof.Rows[j + 1][Col_sta]);

                                                            if (index_sta1 == -1 && sta1 >= sta_start && sta1 <= sta_end)
                                                            {
                                                                index_sta1 = j;
                                                                diferenta1 = sta1 - sta_start;
                                                            }

                                                            if (index_sta1 != -1 )
                                                            {
                                                                j = dt_prof.Rows.Count;
                                                            }

                                                        }
                                                    }
                                                }



                                                double x1 = poly1.StartPoint.X + sta0 + sta1;

                                                if (index_sta1 >= 0)
                                                {
                                                    x1 = poly1.GetPoint2dAt(index_sta1).X + diferenta1;
                                                }

                                                Xline xline1 = new Xline();
                                                xline1.BasePoint = new Point3d(x1, 0, poly1.Elevation);
                                                xline1.SecondPoint = new Point3d(x1, 10, poly1.Elevation);

                                                Point3dCollection col1 = Functions.Intersect_on_both_operands(poly1, xline1);
                                                Point3d inspt1 = new Point3d();
                                                if (col1.Count > 0)
                                                {
                                                    inspt1 = col1[0];
                                                }
                                                else
                                                {
                                                    inspt1 = new Point3d(x1, (ymin + ymax) / 2, 0);
                                                }




                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                string display_sta_string1 = Functions.Get_chainage_from_double(sta1, "m", 1);
                                                


                                                col_atr.Add("STA");
                                                col_val.Add(display_sta_string1);




                                                BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block_name, inspt1, 1, 0, ln1, col_atr, col_val);
                                               



                                            }
                                            Trans1.Commit();
                                        }

                                    }
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
            ThisDrawing.Editor.WriteMessage("\nCommand:");
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }

        private void button_insert_exclusions_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (Functions.IsNumeric(textBox_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_start.Text);
            }

            if (Functions.IsNumeric(textBox_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }
            if (Functions.IsNumeric(textBox_start_prof_sta.Text) == false)
            {
                return;
            }

            double sta0 = Convert.ToDouble(textBox_start_prof_sta.Text);
            try
            {


                if (comboBox_ws1.Text != "")
                {
                    string string1 = comboBox_ws1.Text;
                    if (string1.Contains("[") == true && string1.Contains("]") == true)
                    {
                        string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));

                        string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                        if (filename.Length > 0 && sheet_name.Length > 0)
                        {
                            set_enable_false();
                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                            if (W1 != null)
                            {
                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("sta1", typeof(double));
                                dt1.Columns.Add("sta2", typeof(double));

                                List<string> list_col = new List<string>();
                                List<string> list_colxl = new List<string>();
                                list_col.Add("sta1");
                                list_colxl.Add(textBox_sta1.Text);
                                list_col.Add("sta2");
                                list_colxl.Add(textBox_sta2.Text);

                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, list_col,list_colxl);
                                if (dt1.Rows.Count > 0)
                                {

                                   

                                    ObjectId[] Empty_array = null;
                                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                                    Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                    {
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly2;
                                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly2 = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                        Prompt_poly2.MessageForAdding = "\nselect profile polyline:";
                                        Prompt_poly2.SingleOnly = true;
                                        this.WindowState = FormWindowState.Minimized;
                                        Rezultat_poly2 = ThisDrawing.Editor.GetSelection(Prompt_poly2);

                                        if (Rezultat_poly2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            this.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {
                                            BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                            Entity Ent2 = Trans1.GetObject(Rezultat_poly2.Value[0].ObjectId, OpenMode.ForRead) as Entity;
                                            if ((Ent2 is Polyline) == false)
                                            {
                                                MessageBox.Show("the polyline profile is not a polyline\r\n is a " + Ent2.GetType().ToString() + "\r\nOperation aborted");
                                                this.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }
                                            Polyline poly1 = Ent2 as Polyline;
                                            double ymin = -1000000;
                                            double ymax = 1000000;

                                            for (int i = 0; i < poly1.NumberOfVertices - 1; ++i)
                                            {
                                                double y = poly1.GetPointAtParameter(i).Y;
                                                if (i == 0)
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

                                            string ter_ln = "_db_exclusion";
                                            string block_name = "exclusions";
                                            Functions.Creaza_layer(ter_ln, 2, true);

                                            double deltaY = 0;

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                double sta1 = Convert.ToDouble(dt1.Rows[i][0]);
                                                double sta2 = Convert.ToDouble(dt1.Rows[i][1]);


                                                double diferenta1 = 0;
                                                double diferenta2 = 0;



                                                int index_sta1 = -1;
                                                int index_sta2 = -1;
                                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][Col_sta] != DBNull.Value && dt_prof.Rows[j + 1][Col_sta] != DBNull.Value)
                                                        {
                                                            double sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta]);
                                                            if (dt_prof.Rows[j][Col_sta_eq] != DBNull.Value)
                                                            {
                                                                sta_start = Convert.ToDouble(dt_prof.Rows[j][Col_sta_eq]);
                                                            }
                                                            double sta_end = Convert.ToDouble(dt_prof.Rows[j + 1][Col_sta]);

                                                            if (index_sta1 == -1 && sta1 >= sta_start && sta1 <= sta_end)
                                                            {
                                                                index_sta1 = j;
                                                                diferenta1 = sta1 - sta_start;
                                                            }
                                                            if (index_sta2 == -1 && sta2 >= sta_start && sta2 <= sta_end)
                                                            {
                                                                index_sta2 = j;
                                                                diferenta2 = sta2 - sta_start;
                                                            }
                                                            if (index_sta1 != -1 && index_sta2 != -1)
                                                            {
                                                                j = dt_prof.Rows.Count;
                                                            }

                                                        }
                                                    }
                                                }



                                                double x1 = poly1.StartPoint.X + sta0 + sta1;

                                                if (index_sta1 >= 0)
                                                {
                                                    x1 = poly1.GetPoint2dAt(index_sta1).X + diferenta1;
                                                }

                                                Xline xline1 = new Xline();
                                                xline1.BasePoint = new Point3d(x1, 0, poly1.Elevation);
                                                xline1.SecondPoint = new Point3d(x1, 10, poly1.Elevation);

                                                Point3dCollection col1 = Functions.Intersect_on_both_operands(poly1, xline1);
                                                Point3d inspt1 = new Point3d();
                                                if (col1.Count > 0)
                                                {
                                                    inspt1 = col1[0];
                                                }
                                                else
                                                {
                                                    inspt1 = new Point3d(x1, (ymin + ymax) / 2, 0);
                                                }


                                                double x2 = poly1.StartPoint.X + sta0 + sta2;


                                                if (index_sta2 >= 0)
                                                {
                                                    double x222 = poly1.GetPoint2dAt(index_sta2).X;

                                                    x2 = poly1.GetPoint2dAt(index_sta2).X + diferenta2;
                                                }

                                                Xline xline2 = new Xline();
                                                xline2.BasePoint = new Point3d(x2, 0, poly1.Elevation);
                                                xline2.SecondPoint = new Point3d(x2, 10, poly1.Elevation);

                                                Point3dCollection col2 = Functions.Intersect_on_both_operands(poly1, xline2);
                                                Point3d inspt2 = new Point3d();
                                                if (col2.Count > 0)
                                                {
                                                    inspt2 = col2[0];
                                                }
                                                else
                                                {
                                                    inspt2 = new Point3d(x2, (ymin + ymax) / 2, 0);
                                                }

                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                string display_sta_string1 = Functions.Get_chainage_from_double(sta1, "m", 1);
                                                string display_sta_string2 = Functions.Get_chainage_from_double(sta2, "m", 1);



                                                col_atr.Add("STA1");
                                                col_val.Add(display_sta_string1);

                                                col_atr.Add("STA2");
                                                col_val.Add(display_sta_string2);


                                                BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block_name, inspt1, 1, 0, ter_ln, col_atr, col_val);
                                                double exist_dist2 = Functions.Get_Param_Value_block(block1, "Distance2");
                                          
                                                if (sta2 < sta1)
                                                {
                                                    MessageBox.Show("Sta2 " + sta2.ToString() + " is smaller than \r\nSta1 " + sta1.ToString());
                                                }

                                                double diference = 0;
                                                if (dt_prof != null && dt_prof.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < dt_prof.Rows.Count - 1; ++j)
                                                    {
                                                        if (dt_prof.Rows[j][Col_sta] != DBNull.Value && dt_prof.Rows[j][Col_sta_eq] != DBNull.Value)
                                                        {
                                                            double sta_back = Convert.ToDouble(dt_prof.Rows[j][Col_sta]);

                                                            double sta_ahead = Convert.ToDouble(dt_prof.Rows[j][Col_sta_eq]);
                                                            if (sta2 > sta_ahead && sta1 < sta_back)
                                                            {
                                                                diference = diference + sta_ahead - sta_back;
                                                            }

                                                        }
                                                    }
                                                }

                                                double stretch1 = sta2 - sta1;

                                                Functions.Stretch_block(block1, "Distance1", stretch1 - diference);
                                                Functions.Stretch_block(block1, "Distance2", exist_dist2 + (inspt2.Y - inspt1.Y));
                                             


                                            }
                                            Trans1.Commit();
                                        }

                                    }
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
            ThisDrawing.Editor.WriteMessage("\nCommand:");
            this.WindowState = FormWindowState.Normal;
            set_enable_true();
        }
    }
}
