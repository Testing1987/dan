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
    public partial class Tcal : Form
    {

        bool Freeze_operations = false;
        private bool clickdragdown;
        private Point lastLocation;
        Polyline Poly_avg1 = null;
        List<ObjectId> List_txt = null;
        System.Data.DataTable dt_slope_ranges = null;
        List<ObjectId> List_poly = null;
        string nume_layer;
        double line_width = 1;
        string shindex_excel_name = "sheet_index.xlsx";
        int Start_row_Sheet_index = 11;

        System.Data.DataTable dt_prof = null;
        System.Data.DataTable dt_slope = null;

        public Tcal()
        {
            InitializeComponent();
            nume_layer = "no_plot - [" + Environment.UserName.ToUpper() + "]";
        }

        #region minimize and move close

        private void button_minimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button_Exit_Click(object sender, EventArgs e)
        {
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
                double Dist = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);

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
            string distmin_string = textBox_min.Text;

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

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
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
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }

                            this.WindowState = FormWindowState.Normal;
                            Polyline Poly2d = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead) as Polyline;

                            if (Poly2d == null)
                            {
                                MessageBox.Show("you did not select a polyline");
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }

                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                            dt_slope_ranges = Create_slope_ranges();

                            List_txt = new List<ObjectId>();
                            Functions.Creaza_layer(nume_layer, 30, false);

                            System.Data.DataTable dt1 = creaza_dt_from_poly(Poly2d);

                            #region create profile original labels

                            for (int i = 0; i < Poly2d.NumberOfVertices - 1; ++i)
                            {
                                double x1 = Poly2d.GetPointAtParameter(i).X;
                                double y1 = Poly2d.GetPointAtParameter(i).Y;
                                double x2 = Poly2d.GetPointAtParameter(i + 1).X;
                                double y2 = Poly2d.GetPointAtParameter(i + 1).Y;

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
                            Poly_avg_1_step.AddVertexAt(0, new Point2d(Poly2d.StartPoint.X, Poly2d.StartPoint.Y), 0, 0, 0);

                            Point3d pt1 = Poly2d.StartPoint;
                            int j = 1;
                            double offset1 = Convert.ToDouble(textBox_tolerance.Text) / 2;
                            DBObjectCollection dbcol1 = Poly2d.GetOffsetCurves(offset1);
                            Polyline Poly_down = new Polyline();
                            Poly_down = dbcol1[0] as Polyline;
                            DBObjectCollection dbcol2 = Poly2d.GetOffsetCurves(-offset1);
                            Polyline Poly_up = new Polyline();
                            Poly_up = dbcol2[0] as Polyline;

                            for (int i = 1; i < dt1.Rows.Count; ++i)
                            {

                                Point3d pt2 = Poly2d.GetPointAtParameter(i - 1);
                                Point3d pt3 = Poly2d.GetPointAtParameter(i);

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

                                    Point3d pt4 = Poly2d.GetPointAtParameter(Convert.ToInt32(Poly2d.EndParam));
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
                                    double dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                    if (dist1 < min_dist)
                                    {
                                        dt2.Rows[i]["add_vertex"] = 0;
                                        for (int k = i + 1; k < dt2.Rows.Count - 1; ++k)
                                        {
                                            x2 = Convert.ToDouble(dt2.Rows[k]["x"]);
                                            y2 = Convert.ToDouble(dt2.Rows[k]["y"]);
                                            dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
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
                                                    double dist1 = Math.Pow(Math.Pow(x0 - x, 2) + Math.Pow(y0 - y, 2), 0.5);
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
                    Freeze_operations = false;
                    MessageBox.Show(ex.Message);
                }
            }
            Freeze_operations = false;
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

        private string find__label_from_cat(System.Data.DataTable dt2, double slope1)
        {
            for (int j = 0; j < dt2.Rows.Count; ++j)
            {
                if (dt2.Rows[j]["start"] != DBNull.Value && dt2.Rows[j]["end"] != DBNull.Value)
                {
                    double ss1 = Convert.ToDouble(dt2.Rows[j]["start"]);
                    double se1 = Convert.ToDouble(dt2.Rows[j]["end"]);

                    if (Math.Abs(slope1) >= ss1 && Math.Abs(slope1) < se1)
                    {
                        return Convert.ToString(dt2.Rows[j]["label"]);
                    }
                }
            }
            return "XX";
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

        private void build_poly(int start1, int end1, int colorindex, double width1, string layer_name, Polyline poly2)
        {
            if (checkBox_show_colors_on_profile.Checked == true)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Functions.Creaza_layer(layer_name, 30, false);
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Polyline Poly1 = new Polyline();
                        int j = 0;
                        for (int i = start1; i <= end1; ++i)
                        {
                            Poly1.AddVertexAt(j, poly2.GetPoint2dAt(i), 0, width1, width1);
                            j = j + 1;
                        }

                        Poly1.ColorIndex = colorindex;
                        Poly1.Layer = layer_name;
                        BTrecord.AppendEntity(Poly1);
                        Trans1.AddNewlyCreatedDBObject(Poly1, true);
                        Trans1.Commit();
                        List_poly.Add(Poly1.ObjectId);
                    }
                }
            }
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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                                        Freeze_operations = false;
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
                                            MText existing_label = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                            existing_label.Erase();
                                        }
                                        List_txt.RemoveAt(index1);

                                        LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1);

                                        Point2d pt1 = segm1.StartPoint;
                                        Point2d pt2 = segm1.EndPoint;

                                        double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                        double Texth = 1;
                                        double Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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
                                        Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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
        private void button_remove_vertex_Click(object sender, EventArgs e)
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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

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
                                            MText existing_label2 = Trans1.GetObject(List_txt[index2], OpenMode.ForWrite) as MText;
                                            existing_label2.Erase();

                                        }
                                        List_txt.RemoveAt(index2);

                                        if (List_txt[index1] != ObjectId.Null)
                                        {
                                            MText existing_label1 = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                            existing_label1.Erase();
                                        }
                                        List_txt.RemoveAt(index1);

                                        LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1);

                                        Point2d pt1 = segm1.StartPoint;
                                        Point2d pt2 = segm1.EndPoint;

                                        double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                        double Texth = 1;
                                        double Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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
        private void button_move_vertex_Click(object sender, EventArgs e)
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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                                        Freeze_operations = false;
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
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Point3d new_point = Point_res2.Value;

                                    Poly1.RemoveVertexAt(index1);

                                    Poly1.AddVertexAt(index1, new Point2d(new_point.X, new_point.Y), 0, 0, 0);

                                    if (List_txt[index1] != ObjectId.Null)
                                    {
                                        MText existing_label2 = Trans1.GetObject(List_txt[index1], OpenMode.ForWrite) as MText;
                                        existing_label2.Erase();
                                    }


                                    List_txt.RemoveAt(index1);
                                    if (List_txt[index1 - 1] != ObjectId.Null)
                                    {
                                        MText existing_label1 = Trans1.GetObject(List_txt[index1 - 1], OpenMode.ForWrite) as MText;
                                        existing_label1.Erase();
                                    }

                                    List_txt.RemoveAt(index1 - 1);

                                    LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1 - 1);

                                    Point2d pt1 = segm1.StartPoint;
                                    Point2d pt2 = segm1.EndPoint;

                                    double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    double Texth = 1;
                                    double Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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

                                    Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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
        private void button_remove_mult_vertices_Click(object sender, EventArgs e)
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
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                            Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the optimezed slope polyline:");
                            Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                            Prompt_centerline.AllowNone = true;
                            Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
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
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

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
                                        Freeze_operations = false;
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
                                                MText existing_label2 = Trans1.GetObject(List_txt[i - 1], OpenMode.ForWrite) as MText;
                                                existing_label2.Erase();
                                            }
                                            List_txt.RemoveAt(i - 1);
                                        }
                                        else
                                        {
                                            if (List_txt[i] != ObjectId.Null)
                                            {
                                                MText existing_label2 = Trans1.GetObject(List_txt[i], OpenMode.ForWrite) as MText;
                                                existing_label2.Erase();
                                            }
                                            List_txt.RemoveAt(i);
                                        }
                                    }


                                    if (List_txt[index1 - 1] != ObjectId.Null)
                                    {
                                        MText existing_label3 = Trans1.GetObject(List_txt[index1 - 1], OpenMode.ForWrite) as MText;
                                        existing_label3.Erase();
                                    }
                                    List_txt.RemoveAt(index1 - 1);

                                    LineSegment2d segm1 = Poly1.GetLineSegment2dAt(index1 - 1);

                                    Point2d pt1 = segm1.StartPoint;
                                    Point2d pt2 = segm1.EndPoint;

                                    double Rot1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                    double Texth = 1;
                                    double Dist = Math.Round(pt1.GetDistanceTo(pt2), 1);
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

        public System.Data.DataTable Load_existing_sheet_index(String File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the sheet index data file does not exist");
                return null;
            }

            System.Data.DataTable dt2 = new System.Data.DataTable();

            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                if (Excel1 == null)
                {
                    MessageBox.Show("PROBLEM WITH EXCEL!");
                    return null;
                }

                if (Environment.UserName.ToUpper() == "POP70694")
                {
                    Pgen_mainform.ExcelVisible = true;
                }

                Excel1.Visible = Pgen_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt2 = Functions.Build_Data_table_sheet_index_from_excel(W1, Start_row_Sheet_index + 1);

                    Workbook1.Close();
                    Excel1.Quit();

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt2;

        }


        private void button_export_to_excel_Click(object sender, EventArgs e)
        {

        }
        public System.Data.DataTable Creaza_profile_datatable_structure()
        {



            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add("MMID");
            Lista1.Add("Station");
            Lista1.Add("StationEq");
            Lista1.Add("Elev");
            Lista1.Add("Type");

            Lista2.Add(typeof(string));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(string));


            System.Data.DataTable Data_table_prof = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                Data_table_prof.Columns.Add(Lista1[i], Lista2[i]);
            }
            return Data_table_prof;
        }


        private void button_load_profile_xls_Click(object sender, EventArgs e)
        {
            bool excel_visible = false;
            if (Functions.is_dan_popescu() == true) excel_visible = true;
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {


                    string File1 = fbd.FileName;
                    Microsoft.Office.Interop.Excel.Application Excel1 = null;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }
                    Excel1.Visible = excel_visible;
                    try
                    {
                        Workbook1 = Excel1.Workbooks.Open(File1);
                        W1 = Workbook1.Worksheets[1];




                        dt_prof = Creaza_profile_datatable_structure();


                        Microsoft.Office.Interop.Excel.Range range2 = W1.Range["D9:D30000"];
                        object[,] values2 = new object[30000, 1];
                        values2 = range2.Value2;



                        for (int i = 1; i <= values2.Length; ++i)
                        {
                            object Valoare2 = values2[i, 1];
                            if (Valoare2 != null)
                            {
                                dt_prof.Rows.Add();

                            }
                            else
                            {
                                i = values2.Length + 1;
                            }
                        }

                        int NrR = dt_prof.Rows.Count;
                        int NrC = dt_prof.Columns.Count;




                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[9, 1], W1.Cells[NrR + 8, NrC]];





                        object[,] values = new object[NrR - 1, NrC - 1];

                        values = range1.Value2;

                        for (int i = 0; i < dt_prof.Rows.Count; ++i)
                        {
                            for (int j = 0; j < dt_prof.Columns.Count; ++j)
                            {
                                object Valoare = values[i + 1, j + 1];
                                if (Valoare == null) Valoare = DBNull.Value;

                                dt_prof.Rows[i][j] = Valoare;
                            }
                        }

                        Workbook1.Close();
                        if (Functions.Get_no_of_workbooks_from_Excel() == 0) Excel1.Quit();
                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    }


                }
            }

            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_prof);
        }

        private void button_load_average_slope_Click(object sender, EventArgs e)
        {



            this.WindowState = FormWindowState.Minimized;







            try
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);







                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_poly;
                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_poly = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                Prompt_poly.MessageForAdding = "\nselect optimized slope polyline:";
                Prompt_poly.SingleOnly = true;
                Rezultat_poly = ThisDrawing.Editor.GetSelection(Prompt_poly);

                if (Rezultat_poly.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                {
                    this.MdiParent.WindowState = FormWindowState.Normal;
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

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Polyline poly1 = Ent0 as Polyline;
                        if (poly1 != null)
                        {
                            dt_slope = new System.Data.DataTable();

                            dt_slope.Columns.Add("Station", typeof(double));
                            dt_slope.Columns.Add("Elevation", typeof(double));
                            dt_slope.Columns.Add("Slope", typeof(double));

                            double x0 = poly1.StartPoint.X;
                            double y0 = poly1.StartPoint.Y;

                            double xp = poly1.StartPoint.X;
                            double yp = poly1.StartPoint.Y;

                            for (int i = 0; i < poly1.NumberOfVertices; ++i)
                            {
                                double x1 = poly1.GetPointAtParameter(i).X;
                                double y1 = poly1.GetPointAtParameter(i).Y;
                                dt_slope.Rows.Add();
                                dt_slope.Rows[i]["Station"] = x1 - x0;
                                dt_slope.Rows[i]["Elevation"] = y1 - y0;
                                if (i > 0) dt_slope.Rows[i - 1]["Slope"] = calc_slope(new Point3d(xp, yp, 0), new Point3d(x1, y1, 0));
                                xp = x1;
                                yp = y1;

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


            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_slope);

            this.WindowState = FormWindowState.Normal;
        }

        private void button_export_data_to_excel_Click(object sender, EventArgs e)
        {
            if (dt_slope != null && dt_prof != null && dt_prof.Rows.Count > 0 && dt_slope.Rows.Count > 0)
            {
                System.Data.DataTable dtc = new System.Data.DataTable();

                dtc.Columns.Add("Station", typeof(double));
                dtc.Columns.Add("Elevation", typeof(double));
                dtc.Columns.Add("Average Slope", typeof(double));

                for (int i = 0; i < dt_prof.Rows.Count; ++i)
                {
                    double sta0 = Convert.ToDouble(dt_prof.Rows[i]["Station"]);
                    double elev0 = Convert.ToDouble(dt_prof.Rows[i]["Elev"]);
                    dtc.Rows.Add();
                    dtc.Rows[i]["Station"] = sta0;
                    dtc.Rows[i]["Elevation"] = elev0;
                    double slopep = 0;
                    for (int j = 0; j < dt_slope.Rows.Count; ++j)
                    {
                        if (dt_slope.Rows[j]["Station"] != DBNull.Value)
                        {
                            double sta1 = Convert.ToDouble(dt_slope.Rows[j]["Station"]);
                            double slope1 = 0;

                            if (dt_slope.Rows[j]["Slope"] != DBNull.Value) slope1 = Convert.ToDouble(dt_slope.Rows[j]["Slope"]);
                            if (sta1 > sta0)
                            {
                                dtc.Rows[i]["Average Slope"] = slopep;
                                j = dt_slope.Rows.Count;
                            }
                            if(i== dt_prof.Rows.Count-1) dtc.Rows[i]["Average Slope"] = slopep;
                            slopep = slope1;
                        }

                    }

                }

                Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dtc);
            }
        }
    }
}
