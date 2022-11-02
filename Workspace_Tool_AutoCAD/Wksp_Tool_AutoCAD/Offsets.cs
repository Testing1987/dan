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


        private Polyline get_offset_poly_lod_top(Polyline poly_lod_top, Polyline poly_lod_bottom, double offset_val)
        {



            Point3d pt_on_poly_start = poly_lod_bottom.GetClosestPointTo(poly_lod_top.StartPoint, Vector3d.ZAxis, false);
            Point3d pt_on_poly_end = poly_lod_bottom.GetClosestPointTo(poly_lod_top.EndPoint, Vector3d.ZAxis, false);
            double ds = Math.Pow(Math.Pow(poly_lod_bottom.StartPoint.X - poly_lod_top.StartPoint.X, 2) + Math.Pow(poly_lod_bottom.StartPoint.Y - poly_lod_top.StartPoint.Y, 2), 0.5);
            double de = Math.Pow(Math.Pow(poly_lod_bottom.EndPoint.X - poly_lod_top.EndPoint.X, 2) + Math.Pow(poly_lod_bottom.EndPoint.Y - poly_lod_top.EndPoint.Y, 2), 0.5);


            if ((Math.Round(ds, 3) != Math.Abs(Math.Round(offset_val, 3))) || (Math.Round(de, 3) != Math.Abs(Math.Round(offset_val, 3))))
            {
            start1:
                double sta_start = poly_lod_bottom.GetDistAtPoint(pt_on_poly_start);
                double param1 = poly_lod_bottom.GetParameterAtDistance(sta_start);

                double sta_end = poly_lod_bottom.GetDistAtPoint(pt_on_poly_end);
                double param2 = poly_lod_bottom.GetParameterAtDistance(sta_end);

                Polyline poly1 = new Polyline();


                if (Math.Round(ds, 3) != Math.Abs(Math.Round(offset_val, 3)))
                {
                    int idx1 = 0;
                    for (int i = 0; i < poly_lod_bottom.NumberOfVertices; ++i)
                    {
                        if (i < param1)
                        {
                            poly1.AddVertexAt(idx1, poly_lod_bottom.GetPoint2dAt(i), 0, 0, 0);
                            ++idx1;
                        }
                    }
                    poly1.AddVertexAt(idx1, new Point2d(pt_on_poly_start.X, pt_on_poly_start.Y), 0, 0, 0);
                    DBObjectCollection col_off = poly1.GetOffsetCurves(offset_val);
                    Polyline poly_part1 = col_off[0] as Polyline;


                    Polyline poly_end = new Polyline();
                    poly_end.AddVertexAt(0, poly_part1.GetPoint2dAt(poly_part1.NumberOfVertices - 2), 0, 0, 0);
                    poly_end.AddVertexAt(1, poly_part1.GetPoint2dAt(poly_part1.NumberOfVertices - 1), 0, 0, 0);
                    Polyline poly_start = new Polyline();
                    poly_start.AddVertexAt(0, poly_lod_top.GetPoint2dAt(0), 0, 0, 0);
                    poly_start.AddVertexAt(1, poly_lod_top.GetPoint2dAt(1), 0, 0, 0);

                    Point3dCollection colint1 = new Point3dCollection();

                    poly_end.IntersectWith(poly_start, Intersect.ExtendBoth, colint1, IntPtr.Zero, IntPtr.Zero);

                    if (colint1.Count == 0)
                    {
                        MessageBox.Show("impossible offset - show to dan popescu the linework you have");

                        return null;
                    }

                    int idx = 0;
                    Polyline poly2 = new Polyline();
                    for (int i = 0; i < poly_part1.NumberOfVertices - 1; ++i)
                    {
                        poly2.AddVertexAt(idx, poly_part1.GetPoint2dAt(i), 0, 0, 0);
                        ++idx;
                    }
                    poly2.AddVertexAt(idx, new Point2d(colint1[0].X, colint1[0].Y), 0, 0, 0);
                    ++idx;
                    for (int i = 1; i < poly_lod_top.NumberOfVertices; ++i)
                    {
                        poly2.AddVertexAt(idx, poly_lod_top.GetPoint2dAt(i), 0, 0, 0);
                        ++idx;
                    }
                    poly_lod_top = poly2;
                }
                else if (Math.Round(de, 3) != Math.Abs(Math.Round(offset_val, 3)))
                {

                    int idx1 = 0;
                    poly1.AddVertexAt(idx1, new Point2d(pt_on_poly_end.X, pt_on_poly_end.Y), 0, 0, 0);
                    ++idx1;
                    for (int i = 0; i < poly_lod_bottom.NumberOfVertices; ++i)
                    {
                        if (i > param2)
                        {
                            poly1.AddVertexAt(idx1, poly_lod_bottom.GetPoint2dAt(i), 0, 0, 0);
                            ++idx1;
                        }
                    }

                    DBObjectCollection col_off = poly1.GetOffsetCurves(offset_val);
                    Polyline poly_part2 = col_off[0] as Polyline;

                    Polyline poly_end = new Polyline();
                    poly_end.AddVertexAt(0, poly_part2.GetPoint2dAt(0), 0, 0, 0);
                    poly_end.AddVertexAt(1, poly_part2.GetPoint2dAt(1), 0, 0, 0);
                    Polyline poly_start = new Polyline();
                    poly_start.AddVertexAt(0, poly_lod_top.GetPoint2dAt(poly_lod_top.NumberOfVertices - 2), 0, 0, 0);
                    poly_start.AddVertexAt(1, poly_lod_top.GetPoint2dAt(poly_lod_top.NumberOfVertices - 1), 0, 0, 0);

                    Point3dCollection colint1 = new Point3dCollection();

                    poly_end.IntersectWith(poly_start, Intersect.ExtendBoth, colint1, IntPtr.Zero, IntPtr.Zero);

                    if (colint1.Count == 0)
                    {
                        MessageBox.Show("impossible offset - show to dan popescu the linework you have");
                        return null;
                    }

                    int idx = 0;
                    Polyline poly2 = new Polyline();

                    for (int i = 0; i < poly_lod_top.NumberOfVertices - 1; ++i)
                    {
                        poly2.AddVertexAt(idx, poly_lod_top.GetPoint2dAt(i), 0, 0, 0);
                        ++idx;
                    }
                    poly2.AddVertexAt(idx, new Point2d(colint1[0].X, colint1[0].Y), 0, 0, 0);
                    ++idx;
                    for (int i = 1; i < poly_part2.NumberOfVertices; ++i)
                    {
                        poly2.AddVertexAt(idx, poly_part2.GetPoint2dAt(i), 0, 0, 0);
                        ++idx;
                    }

                    pt_on_poly_start = poly_lod_bottom.GetClosestPointTo(poly2.StartPoint, Vector3d.ZAxis, false);
                    pt_on_poly_end = poly_lod_bottom.GetClosestPointTo(poly2.EndPoint, Vector3d.ZAxis, false);
                    ds = Math.Pow(Math.Pow(poly_lod_bottom.StartPoint.X - poly2.StartPoint.X, 2) + Math.Pow(poly_lod_bottom.StartPoint.Y - poly2.StartPoint.Y, 2), 0.5);
                    de = Math.Pow(Math.Pow(poly_lod_bottom.EndPoint.X - poly2.EndPoint.X, 2) + Math.Pow(poly_lod_bottom.EndPoint.Y - poly2.EndPoint.Y, 2), 0.5);

                    poly_lod_top = poly2;

                    if ((Math.Round(ds, 3) != Math.Abs(Math.Round(offset_val, 3))) || (Math.Round(de, 3) != Math.Abs(Math.Round(offset_val, 3))))
                    {

                        goto start1;
                    }



                }
            }
            return poly_lod_top;
        }
        private Polyline get_atws_top_from_cl(Transaction Trans1, BlockTableRecord BTrecord, Polyline polyCL, Polyline lod3, Point3d pt_start, Point3d pt_end, double width1)
        {
            double sta1 = lod3.GetDistAtPoint(pt_start);
            double sta2 = lod3.GetDistAtPoint(pt_end);
            double par1 = lod3.GetParameterAtDistance(sta1);
            double par2 = lod3.GetParameterAtDistance(sta2);


            Point3d point_on_cl_start = polyCL.GetClosestPointTo(pt_start, Vector3d.ZAxis, false);
            Point3d point_on_cl_end = polyCL.GetClosestPointTo(pt_end, Vector3d.ZAxis, false);
            double sta1_cl = polyCL.GetDistAtPoint(point_on_cl_start);
            double sta2_cl = polyCL.GetDistAtPoint(point_on_cl_end);

            double d_cl1 = Math.Round(Math.Pow(Math.Pow(point_on_cl_start.X - pt_start.X, 2) + Math.Pow(point_on_cl_start.Y - pt_start.Y, 2), 0.5), 3);
            double d_cl2 = Math.Round(Math.Pow(Math.Pow(point_on_cl_end.X - pt_end.X, 2) + Math.Pow(point_on_cl_end.Y - pt_end.Y, 2), 0.5), 3);
            if (d_cl1 != d_cl2)
            {
                Point3d point_on_lod_start = lod3.GetClosestPointTo(point_on_cl_start, Vector3d.ZAxis, false);
                double d_lod11 = Math.Round(Math.Pow(Math.Pow(point_on_cl_start.X - point_on_lod_start.X, 2) + Math.Pow(point_on_cl_start.Y - point_on_lod_start.Y, 2), 0.5), 3);
                Point3d point_on_lod_end = lod3.GetClosestPointTo(point_on_cl_end, Vector3d.ZAxis, false);
                double d_lod21 = Math.Round(Math.Pow(Math.Pow(point_on_cl_end.X - point_on_lod_end.X, 2) + Math.Pow(point_on_cl_end.Y - point_on_lod_end.Y, 2), 0.5), 3);

                if (d_lod11 == d_lod21)
                {
                    d_cl1 = d_cl2;
                }
                else if (d_lod11 == d_cl2)
                {
                    d_cl1 = d_cl2;
                }
                else if (d_cl1 == d_lod21)
                {
                    d_cl1 = d_cl2;
                }

            }



            if (d_cl1 == d_cl2)
            {

                if (width1 < 0) d_cl1 = -d_cl1;
                Polyline atws_top = get_trimmed_offset(polyCL, sta1_cl, sta2_cl, width1 + d_cl1);
                return atws_top;
            }
            else
            {


                Polyline poly1 = get_part_of_poly(lod3, par1, par2);

                Polyline pl1 = new Polyline();
                pl1.AddVertexAt(0, poly1.GetPoint2dAt(0), 0, 0, 0);
                pl1.AddVertexAt(1, poly1.GetPoint2dAt(1), 0, 0, 0);
                pl1.TransformBy(Matrix3d.Scaling(1.1 * width1 * pl1.Length, pl1.StartPoint));
                pl1.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.ZAxis, pl1.StartPoint));


                DBObjectCollection col_offset_lod = poly1.GetOffsetCurves(width1);

                if (col_offset_lod.Count != 1)
                {
                    MessageBox.Show("impossible offset - show the linework to dan popescu #2");

                }

                Polyline atws_top = col_offset_lod[0] as Polyline;

                if (atws_top == null)
                {
                    MessageBox.Show("operation aborted - i can not offset atws_top");

                }

                atws_top = get_offset_poly_lod_top(atws_top, poly1, width1);



                if (atws_top == null)
                {
                    MessageBox.Show("impossible offset - show the linework to dan popescu #3");

                }
                return atws_top;

            }




        }

        private Polyline get_simple_offset_polyline(Polyline poly0, double width1)
        {
           return poly0.GetOffsetCurves(width1)[0] as Polyline;

        }

        private Polyline get_offset_polyline(Polyline poly0, double width1, bool optimize = false)
        {

            Point2dCollection col3 = new Point2dCollection();

            if (poly0 != null)
            {
                if (optimize == true)
                {
                    Point2d pt_prev = new Point2d();
                    double bear_prev = -1000000;
                    for (int i = poly0.NumberOfVertices - 1; i >= 0; --i)
                    {
                        Point2d pt1 = poly0.GetPoint2dAt(i);
                        double d1 = Math.Round(Math.Pow(Math.Pow(pt_prev.X - pt1.X, 2) + Math.Pow(pt_prev.Y - pt1.Y, 2), 0.5), 3);
                        double bear1 = GET_Bearing_rad(pt1.X, pt1.Y, pt_prev.X, pt_prev.Y);

                        if (d1 < 0.001)
                        {
                            poly0.RemoveVertexAt(i);
                        }
                        else if (Math.Round(bear1, 3) == Math.Round(bear_prev, 3))
                        {
                            poly0.RemoveVertexAt(i + 1);
                        }
                        else
                        {
                            pt_prev = pt1;
                            bear_prev = bear1;
                        }

                    }

                }

                bool are_2_vertices = false;



                if (poly0.NumberOfVertices > 2)
                {
                    Point2d pt0 = poly0.GetPoint2dAt(0);
                    for (int i = 0; i < poly0.NumberOfVertices - 2; ++i)
                    {
                        Point2d pt1 = poly0.GetPoint2dAt(i + 1);
                        Point2d pt2 = poly0.GetPoint2dAt(i + 2);
                        double d1 = Math.Round(Math.Pow(Math.Pow(pt0.X - pt1.X, 2) + Math.Pow(pt0.Y - pt1.Y, 2), 0.5), 3);
                        double d2 = Math.Round(Math.Pow(Math.Pow(pt2.X - pt1.X, 2) + Math.Pow(pt2.Y - pt1.Y, 2), 0.5), 3);

                        if (d1 > 0.001 && d2 > 0.001)
                        {
                            Polyline poly1 = new Polyline();
                            poly1.AddVertexAt(0, pt0, 0, 0, 0);
                            poly1.AddVertexAt(1, pt1, 0, 0, 0);
                            Polyline poly2 = new Polyline();
                            poly2.AddVertexAt(0, pt1, 0, 0, 0);
                            poly2.AddVertexAt(1, pt2, 0, 0, 0);

                            DBObjectCollection col1 = poly1.GetOffsetCurves(width1);
                            DBObjectCollection col2 = poly2.GetOffsetCurves(width1);

                            Polyline poly11 = col1[0] as Polyline;
                            if (i == 0) col3.Add(poly11.GetPoint2dAt(0));

                            Polyline poly22 = col2[0] as Polyline;
                            Point3dCollection colint = new Point3dCollection();

                            poly11.IntersectWith(poly22, Intersect.ExtendBoth, colint, IntPtr.Zero, IntPtr.Zero);

                            if (colint.Count > 0)
                            {
                                col3.Add(new Point2d(colint[0].X, colint[0].Y));

                                if (i == poly0.NumberOfVertices - 3) col3.Add(poly22.GetPoint2dAt(1));
                            }
                            pt0 = pt1;
                        }

                    }
                }
                else if (poly0.NumberOfVertices == 2)
                {
                    are_2_vertices = true;
                }
                else
                {
                    MessageBox.Show("length of poly = 0?????");
                }

                if (col3.Count <= 1)
                {
                    col3.Clear();
                    are_2_vertices = true;
                }

                if (are_2_vertices == true)
                {
                    Point2d pt0 = poly0.GetPoint2dAt(0);
                    Point2d pt1 = poly0.GetPoint2dAt(poly0.NumberOfVertices - 1);
                    double d1 = Math.Round(Math.Pow(Math.Pow(pt0.X - pt1.X, 2) + Math.Pow(pt0.Y - pt1.Y, 2), 0.5), 3);
                    if (d1 > 0.001)
                    {
                        Polyline poly1 = new Polyline();
                        poly1.AddVertexAt(0, pt0, 0, 0, 0);
                        poly1.AddVertexAt(1, pt1, 0, 0, 0);
                        DBObjectCollection col1 = poly1.GetOffsetCurves(width1);
                        Polyline poly11 = col1[0] as Polyline;
                        col3.Add(poly11.GetPoint2dAt(0));
                        col3.Add(poly11.GetPoint2dAt(1));
                    }
                }


            }

            Polyline poly3 = null;

            if (col3.Count > 0)
            {
                poly3 = new Polyline();
                int idx = 0;
                Point2d pt_prev = new Point2d();
                for (int i = 0; i < col3.Count; ++i)
                {
                    Point2d pt1 = col3[i];

                    double d1 = Math.Round(Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5), 3);

                    if (d1 > 0.001)
                    {
                        poly3.AddVertexAt(idx, pt1, 0, 0, 0);
                        ++idx;
                        pt_prev = pt1;
                    }

                }
            }

            return poly3;
        }
        static public double GET_Bearing_rad(double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
        }

        private Polyline get_trimmed_offset(Polyline poly1, double sta1, double sta2, double off_val)
        {
            Polyline part_of_cl = null;
            Polyline offset1 = null;
            if (sta1 == sta2 && sta1 == -1)
            {
                offset1 = get_offset_polyline(poly1, off_val);

            }
            else
            {
                double param1 = poly1.GetParameterAtDistance(sta1);
                double param2 = poly1.GetParameterAtDistance(sta2);

                part_of_cl = get_part_of_poly(poly1, param1, param2);
                offset1 = get_offset_polyline(part_of_cl, off_val);
            }



            return offset1;
        }

        private Polyline get_trimmed_offset_with_feature(Polyline poly_cl, double sta1, double sta2, double width_from_cl, double prev_from_cl, double next_from_cl, double wdth)
        {

            bool is_left = false;

            if (wdth < 0) is_left = true;

            Polyline offset1 = get_offset_polyline(poly_cl, width_from_cl);


            Polyline feat1 = null;
            Polyline feat2 = null;



            Point3d pt1 = new Point3d(-1.234, -1.234, 0);
            Point3d pt2 = new Point3d(-1.234, -1.234, 0);

            string col1 = Convert.ToString(Math.Round(sta1, 2));
            string col2 = Convert.ToString(Math.Round(sta2, 2));





            if (dt_sides != null && dt_sides.Rows.Count > 0 && dt_sides.Columns.Contains(col1) == true)
            {
                feat1 = new Polyline();
                int idx = 0;
                for (int k = 0; k < dt_sides.Rows.Count; ++k)
                {
                    if (dt_sides.Rows[k][col1] != DBNull.Value)
                    {
                        feat1.AddVertexAt(idx, (Point2d)dt_sides.Rows[k][col1], 0, 0, 0);
                        ++idx;
                    }
                }


                feat1 = make_polyline_start_point_outside_corridor(poly_cl, offset1, feat1);


                Point3dCollection col_int_feat1 = Functions.Intersect_on_both_operands(feat1, offset1);


                int index_int1 = 0;
                if (col_int_feat1.Count > 0)
                {
                    if (col_int_feat1.Count > 1)
                    {
                        Point2d pt_side = (Point2d)dt_sides.Rows[0][col1 + "#"];
                        double sta_sidef = feat1.GetDistAtPoint(feat1.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat1.Elevation), Vector3d.ZAxis, false));

                        double d1 = 1000;
                        for (int k = 0; k < col_int_feat1.Count; ++k)
                        {
                            Point3d ptt = col_int_feat1[k];

                            double sta2_feat = feat1.GetDistAtPoint(feat1.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                            double d2 = Math.Abs(sta_sidef - sta2_feat);
                            //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                            if (d2 < d1)
                            {
                                d1 = d2;
                                index_int1 = k;
                            }
                        }
                    }
                    pt1 = col_int_feat1[index_int1];
                }
                else
                {
                    double paramcl1 = poly_cl.GetParameterAtDistance(sta1);
                    Polyline part_of_cl1 = get_part_of_poly(poly_cl, paramcl1, poly_cl.EndParam);
                    Polyline offset2 = get_offset_polyline(part_of_cl1, width_from_cl);
                    pt1 = offset2.StartPoint;
                }
            }
            else
            {
                bool is_pi = false;
                double paramcl1 = poly_cl.GetParameterAtDistance(sta1);
                if (Math.Abs(Math.Round(paramcl1, 0) - paramcl1) < 0.01)
                {
                    paramcl1 = Math.Round(paramcl1, 0);
                    if (paramcl1 > 0)
                    {
                        paramcl1 = paramcl1 - 0.5;
                        is_pi = true;
                    }
                }


                Polyline part_of_cl1 = get_part_of_poly(poly_cl, paramcl1, poly_cl.EndParam);
                Polyline offset2 = get_offset_polyline(part_of_cl1, width_from_cl);
                if (is_pi == true)
                {
                    offset2.RemoveVertexAt(0);
                }
                pt1 = offset2.StartPoint;
            }

            if (dt_sides != null && dt_sides.Rows.Count > 0 && dt_sides.Columns.Contains(col2) == true)
            {
                feat2 = new Polyline();
                int idx = 0;
                for (int k = 0; k < dt_sides.Rows.Count; ++k)
                {
                    if (dt_sides.Rows[k][col2] != DBNull.Value)
                    {
                        feat2.AddVertexAt(idx, (Point2d)dt_sides.Rows[k][col2], 0, 0, 0);
                        ++idx;
                    }
                }

                feat2 = make_polyline_start_point_outside_corridor(poly_cl, offset1, feat2);

                Point3dCollection col_int_feat2 = Functions.Intersect_on_both_operands(feat2, offset1);
                int index_int2 = 0;
                if (col_int_feat2.Count > 0)
                {
                    if (col_int_feat2.Count > 1)
                    {


                        Point2d pt_side = (Point2d)dt_sides.Rows[0][col2 + "#"];
                        double sta_sidef = feat2.GetDistAtPoint(feat2.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat2.Elevation), Vector3d.ZAxis, false));
                        double d1 = offset1.Length;
                        for (int k = 0; k < col_int_feat2.Count; ++k)
                        {
                            Point3d ptt = col_int_feat2[k];
                            double sta2_feat = feat2.GetDistAtPoint(feat2.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                            double d2 = Math.Abs(sta_sidef - sta2_feat);
                            //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                            if (d2 < d1)
                            {
                                d1 = d2;
                                index_int2 = k;
                            }
                        }
                    }
                    pt2 = col_int_feat2[index_int2];
                }
                else
                {
                    double paramcl2 = poly_cl.GetParameterAtDistance(sta2);
                    Polyline part_of_cl2 = get_part_of_poly(poly_cl, 0, paramcl2);
                    Polyline offset2 = get_offset_polyline(part_of_cl2, width_from_cl);
                    pt2 = offset2.EndPoint;
                }

            }
            else
            {


                bool is_pi = false;
                double paramcl2 = poly_cl.GetParameterAtDistance(sta2);

                if (poly_cl.EndParam - paramcl2 > 0.1)
                {
                    if (Math.Abs(Math.Round(paramcl2, 0) - paramcl2) < 0.01)
                    {
                        paramcl2 = Math.Round(paramcl2, 0);
                        if (paramcl2 < poly_cl.EndParam)
                        {
                            paramcl2 = paramcl2 + 0.5;
                            is_pi = true;
                        }
                    }
                }

                Polyline part_of_cl2 = get_part_of_poly(poly_cl, 0, paramcl2);
                Polyline offset2 = get_offset_polyline(part_of_cl2, width_from_cl);
                if (is_pi == true)
                {
                    offset2.RemoveVertexAt(offset2.NumberOfVertices - 1);
                }
                pt2 = offset2.EndPoint;
            }

            double param1 = offset1.GetParameterAtPoint(offset1.GetClosestPointTo(pt1, Vector3d.ZAxis, false));
            double param2 = offset1.GetParameterAtPoint(offset1.GetClosestPointTo(pt2, Vector3d.ZAxis, false));

            Polyline poly_offset_result = get_part_of_poly(offset1, param1, param2);



            if (feat1 != null)
            {
                if (width_from_cl != prev_from_cl)
                {
                    double prev_offset = prev_from_cl;
                    if (Math.Abs(prev_from_cl) < Math.Abs(width_from_cl) - Math.Abs(wdth))
                    {
                        prev_offset = width_from_cl - wdth;
                    }

                    Polyline offset_p = get_offset_polyline(poly_cl, prev_offset);

                    Point3dCollection col_int_feat_p = Functions.Intersect_on_both_operands(feat1, offset_p);

                    int index_int1 = 0;
                    if (col_int_feat_p.Count > 0)
                    {
                        if (col_int_feat_p.Count > 1)
                        {
                            Point2d pt_side = (Point2d)dt_sides.Rows[0][col1 + "#"];
                            double sta_sidef = feat1.GetDistAtPoint(feat1.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat1.Elevation), Vector3d.ZAxis, false));
                            double d1 = 1000;
                            for (int k = 0; k < col_int_feat_p.Count; ++k)
                            {
                                Point3d ptt = col_int_feat_p[k];
                                double sta2_feat = feat1.GetDistAtPoint(feat1.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                                double d2 = Math.Abs(sta_sidef - sta2_feat);
                                //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                                if (d2 < d1)
                                {
                                    d1 = d2;
                                    index_int1 = k;
                                }

                            }
                        }

                        Point3d pt_p = col_int_feat_p[index_int1];

                        double par_p = feat1.GetParameterAtPoint(pt_p);
                        double par1 = feat1.GetParameterAtPoint(pt1);
                        Polyline part_feat1 = get_part_of_poly(feat1, par_p, par1);

                        bool against_the_flow = false;
                        if (par_p < par1)
                        {
                            against_the_flow = true;
                        }

                        if (against_the_flow == true)
                        {
                            part_feat1 = reverse_poly(part_feat1);
                        }

                        if (part_feat1.NumberOfVertices > 2)
                        {
                            for (int k = 1; k < part_feat1.NumberOfVertices - 1; ++k)
                            {
                                poly_offset_result.AddVertexAt(0, part_feat1.GetPoint2dAt(k), 0, 0, 0);
                            }
                        }


                    }

                }
            }

            if (feat2 != null)
            {

                if (width_from_cl != next_from_cl)
                {
                    double next_offset = next_from_cl;
                    if (Math.Abs(next_from_cl) < Math.Abs(width_from_cl) - Math.Abs(wdth))
                    {
                        next_offset = width_from_cl - wdth;
                    }

                    Polyline offset_n = get_offset_polyline(poly_cl, next_offset);


                    Point3dCollection col_int_feat_n = Functions.Intersect_on_both_operands(feat2, offset_n);

                    int index_int2 = 0;
                    if (col_int_feat_n.Count > 0)
                    {
                        if (col_int_feat_n.Count > 1)
                        {
                            Point2d pt_side = (Point2d)dt_sides.Rows[0][col2 + "#"];
                            double sta_sidef = feat2.GetDistAtPoint(feat2.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat2.Elevation), Vector3d.ZAxis, false));
                            double d1 = 1000;
                            for (int k = 0; k < col_int_feat_n.Count; ++k)
                            {
                                Point3d ptt = col_int_feat_n[k];
                                double sta2_feat = feat2.GetDistAtPoint(feat2.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                                double d2 = Math.Abs(sta_sidef - sta2_feat);
                                //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                                if (d2 < d1)
                                {
                                    d1 = d2;
                                    index_int2 = k;
                                }
                            }
                        }

                        Point3d pt_n = col_int_feat_n[index_int2];

                        double par_n = feat2.GetParameterAtPoint(pt_n);
                        double par2 = feat2.GetParameterAtPoint(pt2);
                        Polyline part_feat2 = get_part_of_poly(feat2, par_n, par2);

                        bool against_the_flow = false;

                        if (par_n < par2)
                        {
                            against_the_flow = true;
                        }



                        if (against_the_flow == true)
                        {
                            part_feat2 = reverse_poly(part_feat2);
                        }



                        if (part_feat2.NumberOfVertices > 2)
                        {
                            for (int k = 1; k < part_feat2.NumberOfVertices - 1; ++k)
                            {
                                poly_offset_result.AddVertexAt(poly_offset_result.NumberOfVertices, part_feat2.GetPoint2dAt(k), 0, 0, 0);
                            }
                        }


                    }

                }
            }




            return poly_offset_result;

        }

        private Polyline get_trimmed_offset_without_feature(Polyline poly_cl, double sta1, double sta2, double width_from_cl)
        {


            Polyline offset1 = get_offset_polyline(poly_cl, width_from_cl);


            Polyline feat1 = null;
            Polyline feat2 = null;



            Point3d pt1 = new Point3d(-1.234, -1.234, 0);
            Point3d pt2 = new Point3d(-1.234, -1.234, 0);

            string col1 = Convert.ToString(Math.Round(sta1, 2));
            string col2 = Convert.ToString(Math.Round(sta2, 2));





            if (dt_sides != null && dt_sides.Rows.Count > 0 && dt_sides.Columns.Contains(col1) == true)
            {
                feat1 = new Polyline();
                int idx = 0;
                for (int k = 0; k < dt_sides.Rows.Count; ++k)
                {
                    if (dt_sides.Rows[k][col1] != DBNull.Value)
                    {
                        feat1.AddVertexAt(idx, (Point2d)dt_sides.Rows[k][col1], 0, 0, 0);
                        ++idx;
                    }
                }


                feat1 = make_polyline_start_point_outside_corridor(poly_cl, offset1, feat1);


                Point3dCollection col_int_feat1 = Functions.Intersect_on_both_operands(feat1, offset1);


                int index_int1 = 0;
                if (col_int_feat1.Count > 0)
                {
                    if (col_int_feat1.Count > 1)
                    {
                        Point2d pt_side = (Point2d)dt_sides.Rows[0][col1 + "#"];
                        double sta_sidef = feat1.GetDistAtPoint(feat1.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat1.Elevation), Vector3d.ZAxis, false));

                        double d1 = 1000;
                        for (int k = 0; k < col_int_feat1.Count; ++k)
                        {
                            Point3d ptt = col_int_feat1[k];

                            double sta2_feat = feat1.GetDistAtPoint(feat1.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                            double d2 = Math.Abs(sta_sidef - sta2_feat);
                            //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                            if (d2 < d1)
                            {
                                d1 = d2;
                                index_int1 = k;
                            }
                        }
                    }
                    pt1 = col_int_feat1[index_int1];
                }
                else
                {
                    double paramcl1 = poly_cl.GetParameterAtDistance(sta1);
                    Polyline part_of_cl1 = get_part_of_poly(poly_cl, paramcl1, poly_cl.EndParam);
                    Polyline offset2 = get_offset_polyline(part_of_cl1, width_from_cl);
                    pt1 = offset2.StartPoint;
                }
            }
            else
            {
                bool is_pi = false;
                double paramcl1 = poly_cl.GetParameterAtDistance(sta1);
                if (Math.Abs(Math.Round(paramcl1, 0) - paramcl1) < 0.01)
                {
                    paramcl1 = Math.Round(paramcl1, 0);
                    if (paramcl1 > 0)
                    {
                        paramcl1 = paramcl1 - 0.5;
                        is_pi = true;
                    }
                }


                Polyline part_of_cl1 = get_part_of_poly(poly_cl, paramcl1, poly_cl.EndParam);
                Polyline offset2 = get_offset_polyline(part_of_cl1, width_from_cl);
                if (is_pi == true)
                {
                    offset2.RemoveVertexAt(0);
                }
                pt1 = offset2.StartPoint;
            }

            if (dt_sides != null && dt_sides.Rows.Count > 0 && dt_sides.Columns.Contains(col2) == true)
            {
                feat2 = new Polyline();
                int idx = 0;
                for (int k = 0; k < dt_sides.Rows.Count; ++k)
                {
                    if (dt_sides.Rows[k][col2] != DBNull.Value)
                    {
                        feat2.AddVertexAt(idx, (Point2d)dt_sides.Rows[k][col2], 0, 0, 0);
                        ++idx;
                    }
                }

                feat2 = make_polyline_start_point_outside_corridor(poly_cl, offset1, feat2);

                Point3dCollection col_int_feat2 = Functions.Intersect_on_both_operands(feat2, offset1);
                int index_int2 = 0;
                if (col_int_feat2.Count > 0)
                {
                    if (col_int_feat2.Count > 1)
                    {


                        Point2d pt_side = (Point2d)dt_sides.Rows[0][col2 + "#"];
                        double sta_sidef = feat2.GetDistAtPoint(feat2.GetClosestPointTo(new Point3d(pt_side.X, pt_side.Y, feat2.Elevation), Vector3d.ZAxis, false));
                        double d1 = offset1.Length;
                        for (int k = 0; k < col_int_feat2.Count; ++k)
                        {
                            Point3d ptt = col_int_feat2[k];
                            double sta2_feat = feat2.GetDistAtPoint(feat2.GetClosestPointTo(ptt, Vector3d.ZAxis, false));
                            double d2 = Math.Abs(sta_sidef - sta2_feat);
                            //double d2 = Math.Pow(Math.Pow(ptt.X - pt_side.X, 2) + Math.Pow(ptt.Y - pt_side.Y, 2), 0.5);
                            if (d2 < d1)
                            {
                                d1 = d2;
                                index_int2 = k;
                            }
                        }
                    }
                    pt2 = col_int_feat2[index_int2];
                }
                else
                {
                    double paramcl2 = poly_cl.GetParameterAtDistance(sta2);
                    Polyline part_of_cl2 = get_part_of_poly(poly_cl, 0, paramcl2);
                    Polyline offset2 = get_offset_polyline(part_of_cl2, width_from_cl);
                    pt2 = offset2.EndPoint;
                }

            }
            else
            {


                bool is_pi = false;
                double paramcl2 = poly_cl.GetParameterAtDistance(sta2);

                if (poly_cl.EndParam - paramcl2 > 0.1)
                {
                    if (Math.Abs(Math.Round(paramcl2, 0) - paramcl2) < 0.01)
                    {
                        paramcl2 = Math.Round(paramcl2, 0);
                        if (paramcl2 < poly_cl.EndParam)
                        {
                            paramcl2 = paramcl2 + 0.5;
                            is_pi = true;
                        }
                    }
                }

                Polyline part_of_cl2 = get_part_of_poly(poly_cl, 0, paramcl2);
                Polyline offset2 = get_offset_polyline(part_of_cl2, width_from_cl);
                if (is_pi == true)
                {
                    offset2.RemoveVertexAt(offset2.NumberOfVertices - 1);
                }
                pt2 = offset2.EndPoint;
            }

            double param1 = offset1.GetParameterAtPoint(offset1.GetClosestPointTo(pt1, Vector3d.ZAxis, false));
            double param2 = offset1.GetParameterAtPoint(offset1.GetClosestPointTo(pt2, Vector3d.ZAxis, false));

            Polyline poly_offset_result = get_part_of_poly(offset1, param1, param2);


            return poly_offset_result;

        }

        static public void publish_poly(Polyline poly0, int colorindex1)
        {
            if (poly0.NumberOfVertices < 2) return;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Polyline poly1 = new Polyline();

                for (int k = 0; k < poly0.NumberOfVertices; ++k)
                {
                    poly1.AddVertexAt(k, poly0.GetPoint2dAt(k), poly0.GetBulgeAt(k), 0, 0);
                }

                Functions.Creaza_layer("_test", 7, false);

                poly1.Layer = "_test";
                poly1.ColorIndex = colorindex1;

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                BTrecord.AppendEntity(poly1);
                Trans1.AddNewlyCreatedDBObject(poly1, true);
                Trans1.Commit();

            }
        }


        private Polyline get_part_of_poly(Polyline poly0, double par1, double par2, bool optimize = false)
        {
            if (par1 > par2)
            {
                double t = par1;
                par1 = par2;
                par2 = t;
            }

            if (par2 > poly0.EndParam) par2 = poly0.EndParam;

            Polyline poly1 = new Polyline();
            int idx = 0;

            poly1.AddVertexAt(idx, new Point2d(poly0.GetPointAtParameter(par1).X, poly0.GetPointAtParameter(par1).Y), 0, 0, 0);
            ++idx;
            for (int i = 0; i < poly0.NumberOfVertices; ++i)
            {
                if (i > par1 && i < par2)
                {
                    poly1.AddVertexAt(idx, poly0.GetPoint2dAt(i), 0, 0, 0);
                    ++idx;
                }
            }

            poly1.AddVertexAt(idx, new Point2d(poly0.GetPointAtParameter(par2).X, poly0.GetPointAtParameter(par2).Y), 0, 0, 0);

            #region poly optimization 

            if (optimize == true)
            {
                Point2d pt_prev = new Point2d();
                double bear_prev = -1000000;
                for (int i = poly1.NumberOfVertices - 1; i >= 0; --i)
                {
                    Point2d pt1 = poly1.GetPoint2dAt(i);
                    double d1 = Math.Round(Math.Pow(Math.Pow(pt_prev.X - pt1.X, 2) + Math.Pow(pt_prev.Y - pt1.Y, 2), 0.5), 3);
                    double bear1 = GET_Bearing_rad(pt1.X, pt1.Y, pt_prev.X, pt_prev.Y);

                    if (d1 < 0.001)
                    {
                        poly1.RemoveVertexAt(i);
                    }
                    else if (Math.Round(bear1, 3) == Math.Round(bear_prev, 3))
                    {
                        poly1.RemoveVertexAt(i + 1);
                    }
                    else
                    {
                        pt_prev = pt1;
                        bear_prev = bear1;
                    }
                }


            }
            #endregion

            return poly1;
        }

        private Polyline reverse_poly(Polyline poly0)
        {


            Polyline poly1 = new Polyline();

            int idx = 0;
            for (int i = poly0.NumberOfVertices - 1; i >= 0; --i)
            {
                poly1.AddVertexAt(idx, poly0.GetPoint2dAt(i), 0, 0, 0);
                ++idx;
            }


            Point2d pt_prev = new Point2d();
            double bear_prev = -1000000;
            for (int i = poly1.NumberOfVertices - 1; i >= 0; --i)
            {
                Point2d pt1 = poly1.GetPoint2dAt(i);
                double d1 = Math.Round(Math.Pow(Math.Pow(pt_prev.X - pt1.X, 2) + Math.Pow(pt_prev.Y - pt1.Y, 2), 0.5), 3);
                double bear1 = GET_Bearing_rad(pt1.X, pt1.Y, pt_prev.X, pt_prev.Y);

                if (d1 < 0.001)
                {
                    poly1.RemoveVertexAt(i);
                }
                else if (Math.Round(bear1, 3) == Math.Round(bear_prev, 3))
                {
                    poly1.RemoveVertexAt(i + 1);
                }
                else
                {
                    pt_prev = pt1;
                    bear_prev = bear1;
                }

            }
            return poly1;

        }

        public static System.Data.DataTable build_data_table_from_poly(Polyline poly0)
        {

            System.Data.DataTable dt1 = null;


            Polyline poly1 = new Polyline();
            int idx = 0;

            for (int i = 0; i < poly0.NumberOfVertices; ++i)
            {

                poly1.AddVertexAt(idx, poly0.GetPoint2dAt(i), 0, 0, 0);
                ++idx;

            }

            Point2d pt_prev = new Point2d();
            double bear_prev = -1000000;
            for (int i = poly1.NumberOfVertices - 1; i >= 0; --i)
            {
                Point2d pt1 = poly1.GetPoint2dAt(i);
                double d1 = Math.Round(Math.Pow(Math.Pow(pt_prev.X - pt1.X, 2) + Math.Pow(pt_prev.Y - pt1.Y, 2), 0.5), 3);
                double bear1 = GET_Bearing_rad(pt1.X, pt1.Y, pt_prev.X, pt_prev.Y);

                if (d1 < 0.001)
                {
                    poly1.RemoveVertexAt(i);
                }
                else if (Math.Round(bear1, 3) == Math.Round(bear_prev, 3))
                {
                    poly1.RemoveVertexAt(i + 1);
                }
                else
                {
                    pt_prev = pt1;
                    bear_prev = bear1;
                }

            }

            if (poly1.NumberOfVertices > 1 && poly1.Length > 0)
            {
                dt1 = new System.Data.DataTable();
                dt1.Columns.Add("pt", typeof(Point2d));
                dt1.Columns.Add("sta", typeof(double));

                for (int i = 0; i < poly1.NumberOfVertices; ++i)
                {
                    Point3d p1 = poly1.GetPointAtParameter(i);
                    dt1.Rows.Add();
                    dt1.Rows[i][0] = new Point2d(p1.X, p1.Y);
                    dt1.Rows[i][1] = poly1.GetDistanceAtParameter(i);
                }
            }

            return dt1;
        }



    }
}
