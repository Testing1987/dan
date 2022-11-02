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
    public partial class wksp_tool
    {
        private Polyline draw_manual_atws(Polyline polyCL, Polyline lod2, Polyline lod4, Point3d pt_start, Point3d pt_end, double width1, string side1, double orig_sta1, double orig_sta2, string justif1, Polyline edge1, bool edge_at_start)
        {
            double sta1 = lod4.GetDistAtPoint(pt_start);
            double sta2 = lod4.GetDistAtPoint(pt_end);
            double par1 = lod4.GetParameterAtDistance(sta1);
            double par2 = lod4.GetParameterAtDistance(sta2);

            Polyline atws_bottom = get_part_of_poly(lod4, par1, par2);
            Polyline atws_top = get_offset_polyline(atws_bottom, width1);


            Point3d point_on_cl_start = polyCL.GetClosestPointTo(pt_start, Vector3d.ZAxis, false);
            Point3d point_on_cl_end = polyCL.GetClosestPointTo(pt_end, Vector3d.ZAxis, false);
            double sta1_cl = polyCL.GetDistAtPoint(point_on_cl_start);
            double sta2_cl = polyCL.GetDistAtPoint(point_on_cl_end);

            if (orig_sta1 == -1 || orig_sta2 == -1)
            {
                orig_sta1 = sta1_cl;
                orig_sta2 = sta2_cl;
            }



            Point2dCollection col_side1_copy = new Point2dCollection();
            Point2dCollection col_side2_copy = new Point2dCollection();

            if (col_side1 != null && col_side1.Count > 0)
            {
                for (int k = 0; k < col_side1.Count; ++k)
                {
                    col_side1_copy.Add(col_side1[k]);
                }
            }

            if (col_side2 != null && col_side2.Count > 0)
            {
                for (int k = 0; k < col_side2.Count; ++k)
                {
                    col_side2_copy.Add(col_side2[k]);
                }
            }

            #region col_side2
            if (edge1 == null && col_side2 != null && col_side2.Count > 1)
            {
                Point2d pt1 = col_side2[0];

                double d_to_end = Math.Pow(Math.Pow(pt1.X - atws_top.EndPoint.X, 2) + Math.Pow(pt1.Y - atws_top.EndPoint.Y, 2), 0.5);

                Point3d point_on_atws = atws_top.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                double d_to_atws = Math.Pow(Math.Pow(pt1.X - point_on_atws.X, 2) + Math.Pow(pt1.Y - point_on_atws.Y, 2), 0.5);

                if (Math.Round(d_to_end, 3) > 0)
                {
                    if (Math.Round(d_to_atws, 3) > 0)
                    {
                        atws_top.AddVertexAt(atws_top.NumberOfVertices, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                    }
                    else
                    {
                        double par_edge = atws_top.GetParameterAtPoint(point_on_atws);

                        for (int i = atws_top.NumberOfVertices - 1; i >= 0; --i)
                        {
                            if (i > par_edge)
                            {
                                atws_top.RemoveVertexAt(i);
                            }
                        }
                        atws_top.AddVertexAt(atws_top.NumberOfVertices, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                    }
                }


                col_side2.RemoveAt(0);
                col_side2.RemoveAt(col_side2.Count - 1); // this is pt_end


            }
            #endregion

            #region col_side1
            if (edge1 == null && col_side1 != null && col_side1.Count > 1)
            {
                Point2d pt1 = col_side1[col_side1.Count - 1];

                double d_to_start = Math.Pow(Math.Pow(pt1.X - atws_top.StartPoint.X, 2) + Math.Pow(pt1.Y - atws_top.StartPoint.Y, 2), 0.5);

                Point3d point_on_atws = atws_top.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                double d_to_atws = Math.Pow(Math.Pow(pt1.X - point_on_atws.X, 2) + Math.Pow(pt1.Y - point_on_atws.Y, 2), 0.5);

                if (Math.Round(d_to_start, 3) > 0)
                {
                    if (Math.Round(d_to_atws, 3) > 0)
                    {
                        atws_top.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                    }
                    else
                    {
                        double par_edge = atws_top.GetParameterAtPoint(point_on_atws);
                        for (int i = 0; i < atws_top.NumberOfVertices; ++i)
                        {
                            if (i < par_edge)
                            {
                                atws_top.RemoveVertexAt(0);
                            }
                        }
                        atws_top.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                    }
                }

                col_side1.RemoveAt(0);//this is pt_start
                col_side1.RemoveAt(col_side1.Count - 1);


            }
            #endregion



            Polyline poly_atws = new Polyline();
            int idx = 0;

            Point3d point_start = atws_top.StartPoint;

            Point2d lastpt = new Point2d(point_start.X, point_start.Y);
            poly_atws.AddVertexAt(idx, lastpt, 0, 0, 0);
            ++idx;

            for (int i = 1; i < atws_top.NumberOfVertices; ++i)
            {
                Point2d pt1 = atws_top.GetPoint2dAt(i);
                double d0 = Math.Round(Math.Pow(Math.Pow(lastpt.X - pt1.X, 2) + Math.Pow(lastpt.Y - pt1.Y, 2), 0.5), 3);
                if (d0 > 0.001)
                {
                    poly_atws.AddVertexAt(idx, pt1, 0, 0, 0);
                    ++idx;
                    lastpt = pt1;
                }
            }

            if (col_side2 != null && col_side2.Count > 0)
            {
                for (int i = 0; i < col_side2.Count; ++i)
                {
                    Point2d pt1 = col_side2[i];
                    double d0 = Math.Round(Math.Pow(Math.Pow(lastpt.X - pt1.X, 2) + Math.Pow(lastpt.Y - pt1.Y, 2), 0.5), 3);
                    if (d0 > 0.001)
                    {
                        poly_atws.AddVertexAt(idx, pt1, 0, 0, 0);
                        ++idx;
                        lastpt = pt1;
                    }
                }
            }

            for (int i = atws_bottom.NumberOfVertices - 1; i >= 0; --i)
            {
                Point2d pt1 = atws_bottom.GetPoint2dAt(i);
                double d0 = Math.Round(Math.Pow(Math.Pow(lastpt.X - pt1.X, 2) + Math.Pow(lastpt.Y - pt1.Y, 2), 0.5), 3);
                if (d0 > 0.001)
                {
                    poly_atws.AddVertexAt(idx, pt1, 0, 0, 0);
                    ++idx;
                    lastpt = pt1;
                }
            }

            if (col_side1 != null && col_side1.Count > 0)
            {
                for (int i = 0; i < col_side1.Count; ++i)
                {
                    Point2d pt1 = col_side1[i];
                    double d0 = Math.Round(Math.Pow(Math.Pow(lastpt.X - pt1.X, 2) + Math.Pow(lastpt.Y - pt1.Y, 2), 0.5), 3);
                    if (d0 > 0.001)
                    {
                        poly_atws.AddVertexAt(idx, pt1, 0, 0, 0);
                        ++idx;
                        lastpt = pt1;
                    }
                }
            }

            int index_h = 1;
            string handle1 = "temp" + index_h.ToString();
            bool run1 = true;

            if (dt_atws.Rows.Count > 0)
            {
                do
                {
                    bool is_found = false;
                    for (int k = 0; k < dt_atws.Rows.Count; ++k)
                    {
                        if (dt_atws.Rows[k][atws_handle0_column] != DBNull.Value)
                        {
                            string h_existing = Convert.ToString(dt_atws.Rows[k][atws_handle0_column]);
                            if (h_existing == handle1)
                            {
                                ++index_h;
                                handle1 = "temp" + index_h.ToString();
                                is_found = true;
                                k = dt_atws.Rows.Count;
                            }
                        }
                    }
                    if (is_found == false)
                    {
                        run1 = false;
                    }
                } while (run1 == true);
            }

            if (width1 < 0) width1 = -width1;

            dt_atws.Rows.Add();
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_sta1_column] = sta1_cl;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_sta2_column] = sta2_cl;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_sta1_orig_column] = orig_sta1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_sta2_orig_column] = orig_sta2;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_type_column] = atws_regular;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_source_column] = atws_source_manual;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_working_side_column] = side1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_width_column] = width1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_length_column] = Math.Round(sta2 - sta1, 2);
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_justification_column] = justif1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_handle_column] = handle1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_handle0_column] = handle1;
            dt_atws.Rows[dt_atws.Rows.Count - 1][atws_area_column] = poly_atws.Area;

            #region dt_manual_atws

            if (dt_atws_manual == null)
            {
                dt_atws_manual = new System.Data.DataTable();
            }

            if (dt_atws_manual.Columns.Contains(handle1) == false)
            {
                dt_atws_manual.Columns.Add(handle1, typeof(Point2d));
            }
            else
            {
                for (int n = 0; n < dt_atws_manual.Rows.Count; ++n)
                {
                    dt_atws_manual.Rows[n][handle1] = DBNull.Value;
                }
            }

            for (int n = 0; n < poly_atws.NumberOfVertices; ++n)
            {
                if (dt_atws_manual.Rows.Count < n + 1)
                {
                    dt_atws_manual.Rows.Add();
                }
                dt_atws_manual.Rows[n][handle1] = poly_atws.GetPoint2dAt(n);
            }
            #endregion


            #region ABBUTTER calcs
            Point3d pt_abutter1 = atws_bottom.GetPointAtParameter(0);
            Point3d pt_abutter3 = atws_bottom.GetPointAtParameter(atws_bottom.NumberOfVertices - 1);
            Point3d pt_lod1 = lod2.GetClosestPointTo(pt_abutter1, Vector3d.ZAxis, false);
            double d1 = Math.Pow(Math.Pow(pt_abutter1.X - pt_lod1.X, 2) + Math.Pow(pt_abutter1.Y - pt_lod1.Y, 2), 0.5);
            Point3d pt_lod3 = lod2.GetClosestPointTo(pt_abutter3, Vector3d.ZAxis, false);
            double d3 = Math.Pow(Math.Pow(pt_abutter3.X - pt_lod3.X, 2) + Math.Pow(pt_abutter3.Y - pt_lod3.Y, 2), 0.5);


            if (Math.Round(d1, 2) == 0 && Math.Round(d3, 2) == 0)
            {
                dt_atws.Rows[dt_atws.Rows.Count - 1][atws_abutter_column] = "TWS";
            }
            else
            {
                dt_atws.Rows[dt_atws.Rows.Count - 1][atws_abutter_column] = "ATWS";
            }

            #endregion

            #region dt_lod_atws

            if (dt_atws_lod_manual == null)
            {
                dt_atws_lod_manual = new System.Data.DataTable();
            }

            if (dt_atws_lod_manual.Columns.Contains(handle1) == false)
            {
                dt_atws_lod_manual.Columns.Add(handle1, typeof(Point2d));
            }
            else
            {
                for (int n = 0; n < dt_atws_lod_manual.Rows.Count; ++n)
                {
                    dt_atws_lod_manual.Rows[n][handle1] = DBNull.Value;
                }
            }

            if (dt_atws_side1 == null)
            {
                dt_atws_side1 = new System.Data.DataTable();
            }

            if (dt_atws_side1.Columns.Contains(handle1) == false)
            {
                dt_atws_side1.Columns.Add(handle1, typeof(Point2d));
            }
            else
            {
                for (int n = 0; n < dt_atws_side1.Rows.Count; ++n)
                {
                    dt_atws_side1.Rows[n][handle1] = DBNull.Value;
                }
            }

            if (dt_atws_side2 == null)
            {
                dt_atws_side2 = new System.Data.DataTable();
            }

            if (dt_atws_side2.Columns.Contains(handle1) == false)
            {
                dt_atws_side2.Columns.Add(handle1, typeof(Point2d));
            }
            else
            {
                for (int n = 0; n < dt_atws_side2.Rows.Count; ++n)
                {
                    dt_atws_side2.Rows[n][handle1] = DBNull.Value;
                }
            }



            Point2dCollection col_side_edge1 = new Point2dCollection();
            col_side_edge1.Add(atws_bottom.GetPoint2dAt(0));
            col_side_edge1.Add(atws_top.GetPoint2dAt(0));

            Point2dCollection col_side_edge2 = new Point2dCollection();
            col_side_edge2.Add(atws_top.GetPoint2dAt(atws_top.NumberOfVertices - 1));
            col_side_edge2.Add(atws_bottom.GetPoint2dAt(atws_bottom.NumberOfVertices - 1));


            if (edge1 != null)
            {
                if (edge_at_start == true)
                {
                    atws_top.AddVertexAt(0, atws_bottom.GetPoint2dAt(0), 0, 0, 0);
                }
                else
                {
                    atws_top.AddVertexAt(atws_top.NumberOfVertices, atws_bottom.GetPoint2dAt(atws_bottom.NumberOfVertices - 1), 0, 0, 0);
                }
            }
            else
            {
                if (col_side1_copy.Count > 0)
                {
                    for (int k = col_side1_copy.Count - 1; k >= 0; --k)
                    {
                        atws_top.AddVertexAt(0, col_side1_copy[k], 0, 0, 0);
                    }
                }
                if (col_side2_copy.Count > 0)
                {
                    for (int k = 0; k < col_side2_copy.Count; ++k)
                    {
                        atws_top.AddVertexAt(atws_top.NumberOfVertices, col_side2_copy[k], 0, 0, 0);
                    }
                }
            }


            #region dt_atws_lod_manual
            Point2dCollection col1 = new Point2dCollection();

            col1.Add(atws_bottom.GetPoint2dAt(0));

            for (int n = 0; n < atws_top.NumberOfVertices; ++n)
            {
                col1.Add(atws_top.GetPoint2dAt(n));
            }

            col1.Add(atws_bottom.GetPoint2dAt(atws_bottom.NumberOfVertices - 1));

            for (int n = 0; n < col1.Count; ++n)
            {
                if (n == dt_atws_lod_manual.Rows.Count)
                {
                    dt_atws_lod_manual.Rows.Add();
                }
                dt_atws_lod_manual.Rows[n][handle1] = col1[n];
            }
            #endregion

            if (edge1 != null)
            {
                for (int n = 0; n < col_side_edge1.Count; ++n)
                {
                    if (n == dt_atws_side1.Rows.Count)
                    {
                        dt_atws_side1.Rows.Add();
                    }
                    dt_atws_side1.Rows[n][handle1] = col_side_edge1[n];
                }

                for (int n = 0; n < col_side_edge2.Count; ++n)
                {
                    if (n == dt_atws_side2.Rows.Count)
                    {
                        dt_atws_side2.Rows.Add();
                    }
                    dt_atws_side2.Rows[n][handle1] = col_side_edge2[n];
                }
            }
            else
            {
                if (col_side1_copy.Count > 0)
                {
                    for (int n = 0; n < col_side1_copy.Count; ++n)
                    {
                        if (n == dt_atws_side1.Rows.Count)
                        {
                            dt_atws_side1.Rows.Add();
                        }
                        dt_atws_side1.Rows[n][handle1] = col_side1_copy[n];
                    }
                }
                else
                {
                    if (dt_atws_side1.Rows.Count == 0)
                    {
                        dt_atws_side1.Rows.Add();
                    }
                    dt_atws_side1.Rows[0][handle1] = new Point2d(pt_start.X, pt_start.Y);//bottom
                    if (dt_atws_side1.Rows.Count < 2)
                    {
                        dt_atws_side1.Rows.Add();
                    }
                    dt_atws_side1.Rows[1][handle1] = col_side_edge1[1];//top

                }

                if (col_side2_copy.Count > 0)
                {
                    for (int n = 0; n < col_side2_copy.Count; ++n)
                    {
                        if (n == dt_atws_side2.Rows.Count)
                        {
                            dt_atws_side2.Rows.Add();
                        }
                        dt_atws_side2.Rows[n][handle1] = col_side2_copy[n];
                    }
                }
                else
                {
                    if (dt_atws_side2.Rows.Count == 0)
                    {
                        dt_atws_side2.Rows.Add();
                    }
                    dt_atws_side2.Rows[0][handle1] = col_side_edge2[0];//top
                    if (dt_atws_side2.Rows.Count < 2)
                    {
                        dt_atws_side2.Rows.Add();
                    }
                    dt_atws_side2.Rows[1][handle1] = new Point2d(pt_end.X, pt_end.Y);//bottom
                }

            }

            #endregion


            return poly_atws;



        }


        private void build_col_side_1_and_2(Polyline feature1, Polyline lod3, Point3d pt_at_start, Point3d pt_at_end, double width1, bool is_increasing, int side1)
        {
            double param1 = lod3.GetParameterAtPoint(lod3.GetClosestPointTo(pt_at_start, Vector3d.ZAxis, false));
            double param2 = lod3.GetParameterAtPoint(lod3.GetClosestPointTo(pt_at_end, Vector3d.ZAxis, false));

            Polyline part_lod3 = get_part_of_poly(lod3, param1, param2);

            Point3dCollection col_int_feat_bottom = new Point3dCollection();
            feature1.IntersectWith(part_lod3, Intersect.ExtendArgument, col_int_feat_bottom, IntPtr.Zero, IntPtr.Zero);

            if (col_int_feat_bottom.Count == 0)
            {
                MessageBox.Show("feature polyline does not intersect the boundary");
                return;
            }

            int idx_int_bottom = 0;
            if (col_int_feat_bottom.Count > 1)
            {
                double d1 = 1000;
                for (int n = 0; n < col_int_feat_bottom.Count; ++n)
                {
                    Point3d pt2 = col_int_feat_bottom[n];
                    double d2 = 10000000;
                    if (side1 == 1)
                    {
                        d2 = Math.Pow(Math.Pow(pt2.X - pt_at_start.X, 2) + Math.Pow(pt2.Y - pt_at_start.Y, 2), 0.5);
                    }
                    else
                    {
                        d2 = Math.Pow(Math.Pow(pt2.X - pt_at_end.X, 2) + Math.Pow(pt2.Y - pt_at_end.Y, 2), 0.5);
                    }
                    if (d2 < d1)
                    {
                        idx_int_bottom = n;
                        d1 = d2;
                    }
                }
            }


            Polyline lod3_offset = get_offset_polyline(part_lod3, width1);

            Point3dCollection col_int_feat_top = new Point3dCollection();
            feature1.IntersectWith(lod3_offset, Intersect.ExtendArgument, col_int_feat_top, IntPtr.Zero, IntPtr.Zero);


            if (col_int_feat_top.Count == 0)
            {
                MessageBox.Show("feature offset polyline does not intersect the boundary");
                return;
            }

            int idx_int_top = 0;

            if (col_int_feat_top.Count > 1)
            {
                double d1 = 1000;
                for (int n = 0; n < col_int_feat_top.Count; ++n)
                {
                    Point3d pt2 = col_int_feat_top[n];
                    double d2 = 10000;

                    if (side1 == 1)
                    {
                        d2 = Math.Pow(Math.Pow(pt2.X - pt_at_start.X, 2) + Math.Pow(pt2.Y - pt_at_start.Y, 2), 0.5);
                    }
                    else
                    {
                        d2 = Math.Pow(Math.Pow(pt2.X - pt_at_end.X, 2) + Math.Pow(pt2.Y - pt_at_end.Y, 2), 0.5);
                    }

                    if (d2 < d1)
                    {
                        idx_int_top = n;
                        d1 = d2;
                    }
                }
            }

            double param_feat_bottom = feature1.GetParameterAtPoint(col_int_feat_bottom[idx_int_bottom]);
            double param_feat_top = feature1.GetParameterAtPoint(col_int_feat_top[idx_int_top]);

            Polyline part_feat = get_part_of_poly(feature1, param_feat_bottom, param_feat_top);
            double d_f = Math.Pow(Math.Pow(part_feat.StartPoint.X - col_int_feat_top[idx_int_top].X, 2) + Math.Pow(part_feat.StartPoint.Y - col_int_feat_top[idx_int_top].Y, 2), 0.5);

            if (Math.Round(d_f, 3) > 0)
            {
                part_feat = reverse_poly(part_feat);
            }

            if (col_int_feat_top.Count > 1)
            {
                Point3dCollection col_int_feat_top1 = new Point3dCollection();
                part_feat.IntersectWith(lod3_offset, Intersect.ExtendArgument, col_int_feat_top1, IntPtr.Zero, IntPtr.Zero);

                if (col_int_feat_top1.Count > 1)
                {
                    double dist3 = 1000;

                    Point3d pt_int_feat = new Point3d();

                    for (int k = 0; k < col_int_feat_top1.Count; ++k)
                    {
                        Point3d pt3 = col_int_feat_top1[k];
                        double param3 = feature1.GetParameterAtPoint(pt3);
                        if (Math.Abs(param3 - param_feat_bottom) < dist3)
                        {
                            dist3 = Math.Abs(param3 - param_feat_bottom);
                            pt_int_feat = pt3;
                        }
                    }

                    double param_end = part_feat.GetParameterAtPoint(pt_int_feat);
                    if (Math.Abs(param_end - part_feat.EndParam) > 0.1)
                    {
                        part_feat = get_part_of_poly(part_feat, param_end, part_feat.EndParam);
                    }
                }
            }


            if (is_increasing == true)
            {
                col_side1 = new Point2dCollection();
                for (int k = part_feat.NumberOfVertices - 1; k >= 0; --k)
                {
                    col_side1.Add(part_feat.GetPoint2dAt(k));
                }
            }
            else
            {
                col_side2 = new Point2dCollection();
                for (int k = 0; k < part_feat.NumberOfVertices; ++k)
                {
                    col_side2.Add(part_feat.GetPoint2dAt(k));
                }
            }

        }


    }
}
