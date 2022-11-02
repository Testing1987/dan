using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class access_road_methods
    {
        string ar_width_column = "WIDTH";
        string ar_length_column = "LENGTH";
        string ar_type_column = "TYPE";
        string ar_handle_column = "HANDLE";
        string ar_station_column = "STA";
        string atws_handle_column = "HANDLE";

        public System.Data.DataTable get_dt_ar_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add(ar_station_column, typeof(double));
            dt1.Columns.Add(ar_width_column, typeof(double));
            dt1.Columns.Add(ar_length_column, typeof(double));
            dt1.Columns.Add(ar_type_column, typeof(string));
            dt1.Columns.Add(ar_handle_column, typeof(string));

            return dt1;
        }
        public void draw_access_road(Document ThisDrawing, Transaction Trans1,Polyline poly_ar_cl ,System.Data.DataTable dt_cl, System.Data.DataTable dt_ar,
                                        System.Data.DataTable dt_manual_ar, System.Data.DataTable dt_lod_left, System.Data.DataTable dt_lod_right,
                                      ref System.Data.DataTable dt1, ref System.Data.DataTable dt2, ref System.Data.DataTable dt3, ref System.Data.DataTable dt4,
                                        string ar_type, double width1, string ar_layer, RadioButton radioButton_ar_perm, DataGridView dataGridView_ar_data,
                                        System.Windows.Forms.CheckBox checkBox_use_od, ComboBox comboBox_ar_od_name, ComboBox comboBox_ar_od_field)
        {

            List<ObjectId> lista_od_ar_object_id = new List<ObjectId>();
            List<string> lista_od_atws_justif = new List<string>();





                        Polyline lod_left = new Polyline();
                        Polyline lod_right = new Polyline();

                        if (radioButton_ar_perm.Checked == true)
                        {
                            lod_left = wksp_tool.create_lod_construction_polylines(1, "LEFT");
                            lod_right = wksp_tool.create_lod_construction_polylines(1, "RIGHT");
                            if (dt1 == null)
                            {
                                dt1 = new System.Data.DataTable();
                            }
                            if (dt2 == null)
                            {
                                dt2 = new System.Data.DataTable();
                            }
                        }
                        else
                        {
                            lod_left = wksp_tool.create_lod_construction_polylines(4, "LEFT");
                            lod_right = wksp_tool.create_lod_construction_polylines(4, "RIGHT");
                            if (dt3 == null)
                            {
                                dt3 = new System.Data.DataTable();
                            }
                            if (dt4 == null)
                            {
                                dt4 = new System.Data.DataTable();
                            }

                        }


                        Polyline polyCL = new Polyline();
                        for (int i = 0; i < dt_cl.Rows.Count; ++i)
                        {
                            if (dt_cl.Rows[i][0] != DBNull.Value)
                            {
                                polyCL.AddVertexAt(i, (Point2d)dt_cl.Rows[i][0], 0, 0, 0);
                            }
                        }






                      
                        if (poly_ar_cl != null)
                        {
                            poly_ar_cl.Elevation = 0;

                            DBObjectCollection col_off_ar_right = poly_ar_cl.GetOffsetCurves(width1 / 2);
                            DBObjectCollection col_off_ar_left = poly_ar_cl.GetOffsetCurves(-width1 / 2);


                            Polyline p_int_left_side = col_off_ar_left[0] as Polyline;
                            Polyline p_int_right_side = col_off_ar_right[0] as Polyline;

                            bool is_left = false;
                            bool is_right = false;

                            Point3d pt_left_side = new Point3d();
                            Point3d pt_right_side = new Point3d();
                            Point3d pt_center = new Point3d();

                            double ref_sta = 0;

                            if (p_int_left_side != null && p_int_right_side != null)
                            {
                                Point3dCollection col_int_left = Functions.Intersect_on_both_operands(poly_ar_cl, lod_left);
                                Point3dCollection col_int_right = Functions.Intersect_on_both_operands(poly_ar_cl, lod_right);

                                if (col_int_left.Count > 0 && col_int_right.Count > 0)
                                {

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify side of access road");
                                    PP1.AllowNone = true;
                                    Point_res1 = ThisDrawing.Editor.GetPoint(PP1);

                                    if (Point_res1.Status == PromptStatus.OK)
                                    {
                                        Point3d pt_side = Point_res1.Value;

                                        Point3d pt_left = lod_left.GetClosestPointTo(pt_side, Vector3d.ZAxis, false);
                                        Point3d pt_right = lod_right.GetClosestPointTo(pt_side, Vector3d.ZAxis, false);


                                        double d_left = Math.Pow(Math.Pow(pt_side.X - pt_left.X, 2) + Math.Pow(pt_side.Y - pt_left.Y, 2), 0.5);
                                        double d_right = Math.Pow(Math.Pow(pt_side.X - pt_right.X, 2) + Math.Pow(pt_side.Y - pt_right.Y, 2), 0.5);

                                        if (d_left < d_right)
                                        {
                                            is_left = true;

                                            Point3d pt_on_cl = polyCL.GetClosestPointTo(col_int_left[0], Vector3d.ZAxis, false);
                                            ref_sta = polyCL.GetDistAtPoint(pt_on_cl);
                                            pt_center = col_int_left[0];

                                        }
                                        else
                                        {
                                            is_right = true;

                                            Point3d pt_on_cl = polyCL.GetClosestPointTo(col_int_right[0], Vector3d.ZAxis, false);
                                            ref_sta = polyCL.GetDistAtPoint(pt_on_cl);
                                            pt_center = col_int_right[0];

                                        }

                                    }

                                }

                                if (col_int_left.Count > 0 && col_int_right.Count == 0)
                                {
                                    is_left = true;
                                    Point3d pt_on_cl = polyCL.GetClosestPointTo(col_int_left[0], Vector3d.ZAxis, false);
                                    ref_sta = polyCL.GetDistAtPoint(pt_on_cl);
                                    pt_center = col_int_left[0];

                                }

                                if (col_int_right.Count > 0 && col_int_left.Count == 0)
                                {
                                    is_right = true;
                                    Point3d pt_on_cl = polyCL.GetClosestPointTo(col_int_right[0], Vector3d.ZAxis, false);
                                    ref_sta = polyCL.GetDistAtPoint(pt_on_cl);
                                    pt_center = col_int_right[0];

                                }


                                if (is_left == true)
                                {
                                    Point3dCollection col_int_left_side = Functions.Intersect_on_both_operands(p_int_left_side, lod_left);
                                    Point3dCollection col_int_right_side = Functions.Intersect_on_both_operands(p_int_right_side, lod_left);

                                    if (col_int_left_side.Count > 0 && col_int_right_side.Count > 0)
                                    {
                                        pt_left_side = col_int_left_side[0];
                                        pt_right_side = col_int_right_side[0];
                                    }

                                }

                                if (is_right == true)
                                {
                                    Point3dCollection col_int_left_side = Functions.Intersect_on_both_operands(p_int_left_side, lod_right);
                                    Point3dCollection col_int_right_side = Functions.Intersect_on_both_operands(p_int_right_side, lod_right);

                                    if (col_int_left_side.Count > 0 && col_int_right_side.Count > 0)
                                    {
                                        pt_left_side = col_int_left_side[0];
                                        pt_right_side = col_int_right_side[0];
                                    }

                                }

                                if (is_right == false && is_left == false)
                                {
                                    Point3dCollection col_int_cl = new Point3dCollection();
                                    polyCL.IntersectWith(poly_ar_cl, Intersect.ExtendArgument, col_int_cl, IntPtr.Zero, IntPtr.Zero);
                                    if (col_int_cl.Count > 0)
                                    {
                                        Point3d pt_on_cl = polyCL.GetClosestPointTo(col_int_cl[0], Vector3d.ZAxis, false);
                                        ref_sta = polyCL.GetDistAtPoint(pt_on_cl);
                                    }
                                }


                                // start or end of a/r has to be left of left side or right of right side

                                Point3d start_cl_ar = poly_ar_cl.StartPoint;
                                Point3d end_cl_ar = poly_ar_cl.EndPoint;


                                Polyline new_cl_ar = new Polyline();



                                #region is left
                                if (is_left == true)
                                {


                                    double param_int = poly_ar_cl.GetParameterAtPoint(pt_center);

                                    Point3d pt_start = lod_left.GetClosestPointTo(start_cl_ar, Vector3d.ZAxis, false);
                                    Point3d pt_end = lod_left.GetClosestPointTo(end_cl_ar, Vector3d.ZAxis, false);


                                    double d_start = Math.Pow(Math.Pow(start_cl_ar.X - pt_start.X, 2) + Math.Pow(start_cl_ar.Y - pt_start.Y, 2), 0.5);
                                    double d_end = Math.Pow(Math.Pow(end_cl_ar.X - pt_end.X, 2) + Math.Pow(end_cl_ar.Y - pt_end.Y, 2), 0.5);


                                    if (Functions.IsRightDirection(lod_left, start_cl_ar) == false && Functions.IsRightDirection(lod_left, end_cl_ar) == true)
                                    {
                                        int indx1 = 0;

                                        for (int i = 0; i < param_int; ++i)
                                        {
                                            double bulge1 = poly_ar_cl.GetBulgeAt(i);
                                            if (i == Math.Floor(param_int))
                                            {
                                                if (bulge1 != 0)
                                                {
                                                    CircularArc2d circ1 = poly_ar_cl.GetArcSegment2dAt(i);
                                                    double r1 = circ1.Radius;
                                                    double arc_len = poly_ar_cl.GetDistanceAtParameter(param_int) - poly_ar_cl.GetDistanceAtParameter(Math.Floor(param_int));
                                                    double delta1 = arc_len / r1;
                                                    double bulge2 = Math.Tan(delta1 / 4);
                                                    if (bulge1 < 0)
                                                    {
                                                        bulge2 = -bulge2;
                                                    }
                                                    bulge1 = bulge2;
                                                }
                                            }
                                            new_cl_ar.AddVertexAt(indx1, poly_ar_cl.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++indx1;
                                        }

                                        new_cl_ar.AddVertexAt(indx1, new Point2d(pt_center.X, pt_center.Y), 0, 0, 0);

                                        Point3d t = pt_left_side;
                                        pt_left_side = pt_right_side;
                                        pt_right_side = t;
                                    }

                                    if (Functions.IsRightDirection(lod_left, start_cl_ar) == true && Functions.IsRightDirection(lod_left, end_cl_ar) == false)
                                    {
                                        int indx1 = 0;

                                        for (int i = poly_ar_cl.NumberOfVertices - 1; i > param_int; --i)
                                        {
                                            int idx_1 = i - 1;
                                            double bulge1 = 0;
                                            if (idx_1 >= 0)
                                            {
                                                bulge1 = -poly_ar_cl.GetBulgeAt(idx_1);
                                            }

                                            if (i == Math.Ceiling(param_int))
                                            {
                                                if (bulge1 != 0)
                                                {
                                                    CircularArc2d circ1 = poly_ar_cl.GetArcSegment2dAt(idx_1);
                                                    double r1 = circ1.Radius;
                                                    double arc_len = -poly_ar_cl.GetDistanceAtParameter(param_int) + poly_ar_cl.GetDistanceAtParameter(Math.Ceiling(param_int));
                                                    double delta1 = arc_len / r1;
                                                    double bulge2 = Math.Tan(delta1 / 4);
                                                    if (bulge1 < 0)
                                                    {
                                                        bulge2 = -bulge2;
                                                    }
                                                    bulge1 = bulge2;
                                                }

                                            }

                                            new_cl_ar.AddVertexAt(indx1, poly_ar_cl.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++indx1;
                                        }

                                        new_cl_ar.AddVertexAt(indx1, new Point2d(pt_center.X, pt_center.Y), 0, 0, 0);


                                    }


                                    if (new_cl_ar.NumberOfVertices > 1)
                                    {
                                        Polyline left_side = new Polyline();
                                        Polyline right_side = new Polyline();

                                        col_off_ar_right = new_cl_ar.GetOffsetCurves(width1 / 2);
                                        col_off_ar_left = new_cl_ar.GetOffsetCurves(-width1 / 2);
                                        double len1 = Math.Round(new_cl_ar.Length, 2);

                                        p_int_left_side = col_off_ar_left[0] as Polyline;
                                        p_int_right_side = col_off_ar_right[0] as Polyline;



                                        int idxx1 = 0;

                                        for (int i = 0; i < p_int_right_side.NumberOfVertices - 1; ++i)
                                        {
                                            left_side.AddVertexAt(idxx1, p_int_right_side.GetPoint2dAt(i), p_int_right_side.GetBulgeAt(i), 0, 0);
                                            ++idxx1;
                                        }

                                        left_side.AddVertexAt(idxx1, new Point2d(pt_left_side.X, pt_left_side.Y), 0, 0, 0);
                                        idxx1 = 0;

                                        right_side.AddVertexAt(idxx1, new Point2d(pt_right_side.X, pt_right_side.Y), -p_int_left_side.GetBulgeAt(p_int_left_side.NumberOfVertices - 1), 0, 0);
                                        ++idxx1;

                                        for (int i = p_int_left_side.NumberOfVertices - 2; i >= 0; --i)
                                        {
                                            double bulge1 = 0;
                                            if (i - 1 >= 0)
                                            {
                                                bulge1 = -p_int_left_side.GetBulgeAt(i - 1);
                                            }

                                            right_side.AddVertexAt(idxx1, p_int_left_side.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++idxx1;
                                        }

                                        wksp_tool.publish_poly(left_side, 4);
                                        wksp_tool.publish_poly(right_side, 5);


                                        if (radioButton_ar_perm.Checked == true)
                                        {
                                            string col_name_a = Convert.ToString(Math.Round(ref_sta, 2) + "a");
                                            if (dt1.Columns.Contains(col_name_a) == true)
                                            {
                                                dt1.Columns.Remove(col_name_a);
                                            }
                                            dt1.Columns.Add(col_name_a, typeof(Point2d));

                                            string col_name_b = Convert.ToString(Math.Round(ref_sta, 2) + "b");
                                            if (dt1.Columns.Contains(col_name_b) == true)
                                            {
                                                dt1.Columns.Remove(col_name_b);
                                            }
                                            dt1.Columns.Add(col_name_b, typeof(Point2d));

                                            string col_name_c = Convert.ToString(Math.Round(ref_sta, 2) + "wdth_len");
                                            if (dt1.Columns.Contains(col_name_c) == true)
                                            {
                                                dt1.Columns.Remove(col_name_c);
                                            }
                                            dt1.Columns.Add(col_name_c, typeof(double));

                                            for (int i = 0; i < left_side.NumberOfVertices; ++i)
                                            {
                                                if (dt1.Rows.Count == i)
                                                {
                                                    dt1.Rows.Add();
                                                }
                                                dt1.Rows[i][col_name_a] = left_side.GetPoint2dAt(i);
                                            }
                                            for (int i = 0; i < right_side.NumberOfVertices; ++i)
                                            {
                                                if (dt1.Rows.Count == i)
                                                {
                                                    dt1.Rows.Add();
                                                }
                                                dt1.Rows[i][col_name_b] = right_side.GetPoint2dAt(i);
                                            }
                                            dt1.Rows[0][col_name_c] = len1;
                                            dt1.Rows[1][col_name_c] = width1;

                                        }
                                        else
                                        {
                                            string col_name_a = Convert.ToString(Math.Round(ref_sta, 2) + "a");
                                            if (dt3.Columns.Contains(col_name_a) == true)
                                            {
                                                dt3.Columns.Remove(col_name_a);
                                            }
                                            dt3.Columns.Add(col_name_a, typeof(Point2d));

                                            string col_name_b = Convert.ToString(Math.Round(ref_sta, 2) + "b");
                                            if (dt3.Columns.Contains(col_name_b) == true)
                                            {
                                                dt3.Columns.Remove(col_name_b);
                                            }
                                            dt3.Columns.Add(col_name_b, typeof(Point2d));

                                            string col_name_c = Convert.ToString(Math.Round(ref_sta, 2) + "wdth_len");
                                            if (dt3.Columns.Contains(col_name_c) == true)
                                            {
                                                dt3.Columns.Remove(col_name_c);
                                            }
                                            dt3.Columns.Add(col_name_c, typeof(double));


                                            for (int i = 0; i < left_side.NumberOfVertices; ++i)
                                            {
                                                if (dt3.Rows.Count == i)
                                                {
                                                    dt3.Rows.Add();
                                                }
                                                dt3.Rows[i][col_name_a] = left_side.GetPoint2dAt(i);
                                            }
                                            for (int i = 0; i < right_side.NumberOfVertices; ++i)
                                            {
                                                if (dt3.Rows.Count == i)
                                                {
                                                    dt3.Rows.Add();
                                                }
                                                dt3.Rows[i][col_name_b] = right_side.GetPoint2dAt(i);
                                            }


                                            dt1.Rows[0][col_name_c] = len1;
                                            dt1.Rows[1][col_name_c] = width1;
                                        }

                                        if (dt_ar == null)
                                        {
                                            dt_ar = get_dt_ar_structure();
                                        }

                                        dt_ar.Rows.Add();
                                        dt_ar.Rows[dt_ar.Rows.Count - 1][ar_width_column] = width1;
                                        dt_ar.Rows[dt_ar.Rows.Count - 1][ar_type_column] = "Permanent";
                                        dt_ar.Rows[dt_ar.Rows.Count - 1][ar_handle_column] = Convert.ToString(Math.Round(ref_sta, 2));
                                        dt_ar.Rows[dt_ar.Rows.Count - 1][ar_station_column] = Math.Round(ref_sta, 2);
                                        dt_ar.Rows[dt_ar.Rows.Count - 1][ar_length_column] = Math.Round(len1, 2);

                                        wksp_tool.form1.delete_existing_linework();
                                        wksp_tool.form1.draw_all_corridors();

                                    }

                                }
                                #endregion

                                #region is right
                                if (is_right == true)
                                {


                                    double param_int = poly_ar_cl.GetParameterAtPoint(pt_center);

                                    Point3d pt_start = lod_right.GetClosestPointTo(start_cl_ar, Vector3d.ZAxis, false);
                                    Point3d pt_end = lod_right.GetClosestPointTo(end_cl_ar, Vector3d.ZAxis, false);


                                    double d_start = Math.Pow(Math.Pow(start_cl_ar.X - pt_start.X, 2) + Math.Pow(start_cl_ar.Y - pt_start.Y, 2), 0.5);
                                    double d_end = Math.Pow(Math.Pow(end_cl_ar.X - pt_end.X, 2) + Math.Pow(end_cl_ar.Y - pt_end.Y, 2), 0.5);


                                    if (Functions.IsRightDirection(lod_right, start_cl_ar) == true && Functions.IsRightDirection(lod_right, end_cl_ar) == false)
                                    {
                                        int indx1 = 0;

                                        for (int i = 0; i < param_int; ++i)
                                        {

                                            double bulge1 = poly_ar_cl.GetBulgeAt(i);

                                            if (i == Math.Floor(param_int))
                                            {
                                                if (bulge1 != 0)
                                                {
                                                    CircularArc2d circ1 = poly_ar_cl.GetArcSegment2dAt(i);
                                                    double r1 = circ1.Radius;
                                                    double arc_len = poly_ar_cl.GetDistanceAtParameter(param_int) - poly_ar_cl.GetDistanceAtParameter(Math.Floor(param_int));
                                                    double delta1 = arc_len / r1;
                                                    double bulge2 = Math.Tan(delta1 / 4);
                                                    if (bulge1 < 0)
                                                    {
                                                        bulge2 = -bulge2;
                                                    }
                                                    bulge1 = bulge2;
                                                }

                                            }

                                            new_cl_ar.AddVertexAt(indx1, poly_ar_cl.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++indx1;
                                        }

                                        new_cl_ar.AddVertexAt(indx1, new Point2d(pt_center.X, pt_center.Y), 0, 0, 0);



                                        Point3d t = pt_left_side;
                                        pt_left_side = pt_right_side;
                                        pt_right_side = t;

                                    }

                                    if (Functions.IsRightDirection(lod_right, start_cl_ar) == false && Functions.IsRightDirection(lod_right, end_cl_ar) == true)
                                    {
                                        int indx1 = 0;

                                        for (int i = poly_ar_cl.NumberOfVertices - 1; i > param_int; --i)
                                        {
                                            int idx_1 = i - 1;
                                            double bulge1 = 0;
                                            if (idx_1 >= 0)
                                            {
                                                bulge1 = -poly_ar_cl.GetBulgeAt(idx_1);
                                            }

                                            if (i == Math.Ceiling(param_int))
                                            {
                                                if (bulge1 != 0)
                                                {
                                                    CircularArc2d circ1 = poly_ar_cl.GetArcSegment2dAt(idx_1);
                                                    double r1 = circ1.Radius;
                                                    double arc_len = -poly_ar_cl.GetDistanceAtParameter(param_int) + poly_ar_cl.GetDistanceAtParameter(Math.Ceiling(param_int));
                                                    double delta1 = arc_len / r1;
                                                    double bulge2 = Math.Tan(delta1 / 4);
                                                    if (bulge1 < 0)
                                                    {
                                                        bulge2 = -bulge2;
                                                    }
                                                    bulge1 = bulge2;
                                                }

                                            }

                                            new_cl_ar.AddVertexAt(indx1, poly_ar_cl.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++indx1;
                                        }

                                        new_cl_ar.AddVertexAt(indx1, new Point2d(pt_center.X, pt_center.Y), 0, 0, 0);

                                    }

                                    if (new_cl_ar.NumberOfVertices > 1)
                                    {
                                        Polyline left_side = new Polyline();
                                        Polyline right_side = new Polyline();

                                        double len1 = Math.Round(new_cl_ar.Length, 2);

                                        col_off_ar_right = new_cl_ar.GetOffsetCurves(width1 / 2);
                                        col_off_ar_left = new_cl_ar.GetOffsetCurves(-width1 / 2);


                                        p_int_left_side = col_off_ar_left[0] as Polyline;
                                        p_int_right_side = col_off_ar_right[0] as Polyline;



                                        int idxx1 = 0;

                                        for (int i = 0; i < p_int_right_side.NumberOfVertices - 1; ++i)
                                        {
                                            left_side.AddVertexAt(idxx1, p_int_right_side.GetPoint2dAt(i), p_int_right_side.GetBulgeAt(i), 0, 0);
                                            ++idxx1;
                                        }

                                        left_side.AddVertexAt(idxx1, new Point2d(pt_left_side.X, pt_left_side.Y), 0, 0, 0);
                                        idxx1 = 0;

                                        right_side.AddVertexAt(idxx1, new Point2d(pt_right_side.X, pt_right_side.Y), -p_int_left_side.GetBulgeAt(p_int_left_side.NumberOfVertices - 1), 0, 0);
                                        ++idxx1;

                                        for (int i = p_int_left_side.NumberOfVertices - 2; i >= 0; --i)
                                        {
                                            double bulge1 = 0;
                                            if (i - 1 >= 0)
                                            {
                                                bulge1 = -p_int_left_side.GetBulgeAt(i - 1);
                                            }
                                            right_side.AddVertexAt(idxx1, p_int_left_side.GetPoint2dAt(i), bulge1, 0, 0);
                                            ++idxx1;
                                        }






                                    }

                                }
                                #endregion


                               

                            }
                        }



                        
                   

               






        }

        public System.Data.DataTable build_dt_ar_from_config_excel(Worksheet W1)
        {
            System.Data.DataTable dt1 = get_dt_ar_structure();
            string Col1 = "E";

            Range range2 = W1.Range[Col1 + "2:" + Col1 + "30002"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    dt1.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }

            if (is_data == false)
            {
                return null;
            }

            int NrR = dt1.Rows.Count;

            Range range1 = W1.Range["A2:E" + Convert.ToString(NrR + 1)];
            object[,] values = new object[NrR, 5];
            values = range1.Value2;

            for (int i = 0; i < dt1.Rows.Count; ++i)
            {
                for (int j = 0; j < dt1.Columns.Count; ++j)
                {
                    object val = values[i + 1, j + 1];
                    if (val == null) val = DBNull.Value;

                    dt1.Rows[i][j] = val;
                }
            }

            return dt1;
        }

        public void button_highlight_ar_Click(DataGridView dataGridView_ar_data)
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        bool ask_for_selection = false;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                        if (Rezultat_object.Status == PromptStatus.OK)
                        {
                            if (Rezultat_object.Value.Count == 0)
                            {
                                ask_for_selection = true;
                            }
                            if (Rezultat_object.Value.Count > 1)
                            {
                                MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                ask_for_selection = false;
                            }
                        }
                        else ask_for_selection = true;



                        if (ask_for_selection == true)
                        {
                            wksp_tool.minimize_form();
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_object.MessageForAdding = "\nSelect an ATWS";
                            Prompt_object.SingleOnly = true;
                            Rezultat_object = Editor1.GetSelection(Prompt_object);

                        }


                        if (Rezultat_object.Status != PromptStatus.OK)
                        {
                            wksp_tool.maximize_form();

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");

                            return;
                        }
                        wksp_tool.maximize_form();


                        Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                        string handle1 = Ent1.ObjectId.Handle.Value.ToString();

                        for (int i = 0; i < dataGridView_ar_data.Rows.Count; ++i)
                        {
                            string handle2 = Convert.ToString(dataGridView_ar_data.Rows[i].Cells[atws_handle_column].Value);
                            if (handle1 == handle2)
                            {
                                dataGridView_ar_data.CurrentCell = dataGridView_ar_data.Rows[i].Cells[0];
                                i = dataGridView_ar_data.Rows.Count;
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

        public void comboBox_ar_od_name_SelectedIndexChanged(ComboBox comboBox_ar_od_name, ComboBox comboBox_ar_od_field)
        {
            try
            {
                Functions.load_object_data_fieds_to_combobox(comboBox_ar_od_name, comboBox_ar_od_field);
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        public void format_and_transfer_dt_ar_to_excel(Worksheet W9, System.Data.DataTable dt_ar)
        {
            if (W9 != null && dt_ar != null && dt_ar.Rows.Count > 0)
            {
                W9.Range["A:E"].ColumnWidth = 15;

                W9.Range["A1:E1"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                W9.Range["A1:E1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                W9.Range["A1:E30000"].ClearContents();
                W9.Range["A2:E" + Convert.ToString(1 + dt_ar.Rows.Count)].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                W9.Range["A2:E" + Convert.ToString(1 + dt_ar.Rows.Count)].VerticalAlignment = XlVAlign.xlVAlignCenter;
                Range range1 = W9.Range["A1:E1"];
                Functions.Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;

                Functions.Transfer_datatable_to_excel_spreadsheet(W9, dt_ar, 1, true);
                range1 = W9.Range["A2:C" + Convert.ToString(dt_ar.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 44); //orange
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
                range1 = W9.Range["D2:D" + Convert.ToString(dt_ar.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 43); //light green
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;
                range1 = W9.Range["E2:E" + Convert.ToString(dt_ar.Rows.Count + 1)];
                Functions.Color_border_range_inside(range1, 44); //orange
                range1.Font.ColorIndex = 1;//black
                range1.Font.Size = 11;
                range1.Font.Bold = true;

            }
            else
            {
                W9.Range["A1:E30000"].ClearContents();
            }
        }


      



        public void comboBox_ar_od_name_DropDown(ComboBox comboBox_ar_od_name, ComboBox comboBox_ar_od_field)
        {
            try
            {
                Functions.load_object_data_table_name_to_combobox(comboBox_ar_od_name);
                if (comboBox_ar_od_name.Items.Count == 1)
                {
                    comboBox_ar_od_field.Items.Clear();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void attach_od_to_ar(List<ObjectId> lista_od_ar_object_id, List<string> lista_od_ar_justif, System.Windows.Forms.CheckBox checkBox_use_od, ComboBox comboBox_ar_od_name, ComboBox comboBox_ar_od_field)
        {
            if (checkBox_use_od.Checked == false) return;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
          using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                if (lista_od_ar_object_id.Count > 0)
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;



                    List<string> lista_field_name = Functions.get_object_data_table_field_names(Tables1, comboBox_ar_od_name.Text);

                    if (lista_field_name == null || lista_field_name.Count == 0 || lista_field_name.Contains(comboBox_ar_od_field.Text) == false)
                    {
                        MessageBox.Show("Issue with ATWS data table found in the drawing");
                        return;
                    }

                    List<Autodesk.Gis.Map.Constants.DataType> lista_types = Functions.get_object_data_table_data_types(Tables1, comboBox_ar_od_name.Text);

                    for (int i = 0; i < lista_od_ar_object_id.Count; ++i)
                    {

                        Polyline atws1 = Trans1.GetObject(lista_od_ar_object_id[i], OpenMode.ForWrite) as Polyline;

                        List<object> lista_val = new List<object>();

                        for (int k = 0; k < lista_field_name.Count; ++k)
                        {
                            if (lista_field_name[k] == comboBox_ar_od_field.Text)
                            {
                                if (lista_types[k] == Autodesk.Gis.Map.Constants.DataType.Character)
                                {
                                    lista_val.Add(lista_od_ar_justif[i]);
                                }
                                else
                                {
                                    MessageBox.Show("The Object Data field " + comboBox_ar_od_field.Text + " is not defined as character field.\r\nPlease make sure you selected the correct field.\r\nOperation aborted");
                                    Entity ent1 = Trans1.GetObject(lista_od_ar_object_id[i], OpenMode.ForWrite) as Entity;
                                    ent1.Erase();
                                    Trans1.Commit();
                                    return;
                                }
                            }
                            else
                            {
                                lista_val.Add(null);
                            }
                        }
                        Functions.Populate_object_data_table_from_objectid(lista_od_ar_object_id[i], comboBox_ar_od_name.Text, lista_val, lista_types);
                    }
                }
                Trans1.Commit();
            }
        }

        public void attach_dt_ar_to_datagridview(System.Data.DataTable dt1, DataGridView dataGridView_ar_data)
        {
            if (dt1 != null && dt1.Rows.Count > 0)
            {
                dataGridView_ar_data.DataSource = dt1;
                dataGridView_ar_data.Columns[ar_station_column].Width = 75;
                dataGridView_ar_data.Columns[ar_width_column].Width = 50;
                dataGridView_ar_data.Columns[ar_length_column].Width = 60;
                dataGridView_ar_data.Columns[ar_type_column].Width = 150;
                dataGridView_ar_data.Columns[ar_type_column].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                dataGridView_ar_data.Columns[ar_handle_column].Width = 150;

                dataGridView_ar_data.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_ar_data.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_ar_data.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                Padding newpadding = new Padding(4, 0, 0, 0);
                dataGridView_ar_data.ColumnHeadersDefaultCellStyle.Padding = newpadding;
                dataGridView_ar_data.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_ar_data.DefaultCellStyle.BackColor = Color.FromArgb(51, 51, 55);
                dataGridView_ar_data.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_ar_data.EnableHeadersVisualStyles = false;
            }
            else
            {
                dataGridView_ar_data.DataSource = null;
            }
        }






        public void button_zoom_to_ar_Click(System.Data.DataTable dt_ar, DataGridView dataGridView_ar_data)
        {
            try
            {
                if (dt_ar != null && dt_ar.Rows.Count > 0)
                {

                    int row_idx = dataGridView_ar_data.SelectedCells[0].RowIndex;
                    if (row_idx >= 0)
                    {
                        string handle1 = Convert.ToString(dataGridView_ar_data.Rows[row_idx].Cells[ar_handle_column].Value);

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    Functions.zoom_to_object(id1);
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

        public void button_ar_update_justification_Click(System.Data.DataTable dt_ar, System.Windows.Forms.CheckBox checkBox_use_od, ComboBox comboBox_ar_od_name, ComboBox comboBox_ar_od_field)
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        List<ObjectId> lista_od_ar_object_id = new List<ObjectId>();
                        List<string> lista_od_ar_justif = new List<string>();



                        for (int i = 0; i < dt_ar.Rows.Count; ++i)
                        {
                            if (dt_ar.Rows[i][ar_handle_column] != DBNull.Value && dt_ar.Rows[i][ar_type_column] != DBNull.Value)
                            {
                                string handle1 = Convert.ToString(dt_ar.Rows[i][ar_handle_column]);
                                string justif1 = Convert.ToString(dt_ar.Rows[i][ar_type_column]);

                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);

                                if (id1 != ObjectId.Null)
                                {
                                    lista_od_ar_justif.Add(justif1);
                                    lista_od_ar_object_id.Add(id1);
                                }
                            }
                        }

                        attach_od_to_ar( lista_od_ar_object_id, lista_od_ar_justif, checkBox_use_od, comboBox_ar_od_name, comboBox_ar_od_field);

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

        public void button_ar_out_Click(System.Data.DataTable dt_ar, string ar_data)
        {

            try
            {
                if (dt_ar != null && dt_ar.Rows.Count > 0)
                {
                    Worksheet W1 = Functions.get_worksheet_W1(true, ar_data);
                    format_and_transfer_dt_ar_to_excel(W1, dt_ar);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void button_ar_in_Click(System.Data.DataTable dt_ar, string ar_data, DataGridView dataGridView_ar_data)
        {
            try
            {

                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                bool is_found = false;
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    if (Excel1 != null)
                    {

                        foreach (Workbook workbook1 in Excel1.Workbooks)
                        {
                            if (is_found == false)
                            {
                                foreach (Worksheet W1 in workbook1.Worksheets)
                                {
                                    if (is_found == false && W1.Name == ar_data)
                                    {
                                        is_found = true;
                                        dt_ar = build_dt_ar_from_config_excel(W1);
                                        attach_dt_ar_to_datagridview(dt_ar, dataGridView_ar_data);
                                    }
                                }
                            }

                        }
                    }



                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("no excel found");

                }


            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}
