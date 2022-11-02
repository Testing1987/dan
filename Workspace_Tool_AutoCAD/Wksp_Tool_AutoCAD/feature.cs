using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;

namespace Alignment_mdi
{
    public partial class wksp_tool
    {

        private Point3d get_feature_intersection_point(Transaction Trans1, Polyline poly_cl)
        {
            Point3d pt1 = new Point3d(-1.234, -1.234, 0);
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_feat1;
            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_feat1;
            Prompt_feat1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the feature polyline:");
            Prompt_feat1.SetRejectMessage("\nSelect a polyline!");
            Prompt_feat1.AllowNone = true;
            Prompt_feat1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
            Rezultat_feat1 = ThisDrawing.Editor.GetEntity(Prompt_feat1);

            if (Rezultat_feat1.Status != PromptStatus.OK)
            {
                return pt1;
            }

            Polyline feature1 = Trans1.GetObject(Rezultat_feat1.ObjectId, OpenMode.ForRead) as Polyline;
            if (feature1 != null)
            {
                System.Data.DataTable dt_feat_side = build_data_table_from_poly(feature1);

                Point3dCollection col_int_feature = Functions.Intersect_on_both_operands(poly_cl, feature1);
                if (col_int_feature.Count == 1)
                {
                    pt1 = col_int_feature[0];
                }
                else
                {
                    if (col_int_feature.Count > 0)
                    {
                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_feat1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify the feature point");
                        PP1.AllowNone = false;
                        Point_res_feat1 = Editor1.GetPoint(PP1);

                        if (Point_res_feat1.Status != PromptStatus.OK)
                        {
                            return pt1;
                        }
                        Point3d pp1_feat = Point_res_feat1.Value;
                        double d1 = 1000;
                        for (int n = 0; n < col_int_feature.Count; ++n)
                        {
                            Point3d pp2 = col_int_feature[n];
                            double d2 = Math.Pow(Math.Pow(pp1_feat.X - pp2.X, 2) + Math.Pow(pp1_feat.Y - pp2.Y, 2), 0.5);
                            if (d2 < d1)
                            {
                                d1 = d2;
                                pt1 = pp2;
                            }
                        }
                    }
                }

                string col1 = Convert.ToString(Math.Round(poly_cl.GetDistAtPoint(pt1), 2));
                if (dt_sides == null)
                {
                    dt_sides = new System.Data.DataTable();
                }

                if (dt_sides.Columns.Contains(col1) == false)
                {
                    dt_sides.Columns.Add(col1, typeof(Point2d));
                    dt_sides.Columns.Add(col1 + "#", typeof(Point2d));
                }
                else
                {
                    for (int i = 0; i < dt_sides.Rows.Count; ++i)
                    {
                        dt_sides.Rows[i][col1] = DBNull.Value;
                        dt_sides.Rows[i][col1 + "#"] = DBNull.Value;
                    }
                }

                for (int i = 0; i < dt_feat_side.Rows.Count; ++i)
                {
                    if (dt_sides.Rows.Count == i)
                    {
                        dt_sides.Rows.Add();
                    }
                    dt_sides.Rows[i][col1] = dt_feat_side.Rows[i][0];
                }
                dt_sides.Rows[0][col1 + "#"] = new Point2d(pt1.X, pt1.Y);
            }

            return pt1;
        }

        private void set_corridor_offsets(ref double perm_l, ref double tws_l, ref double atws_l, ref double perm_r, ref double tws_r, ref double atws_r, string ws_name, bool flip1)
        {



            for (int j = 0; j < dt_library.Rows.Count; ++j)
            {
                if (dt_library.Rows[j][lib_name_column] != DBNull.Value)
                {
                    string nume1 = Convert.ToString(dt_library.Rows[j][lib_name_column]);
                    if (nume1 == ws_name)
                    {
                        if (dt_library.Rows[j][wksp_atws_left_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_atws_left_column])) == true)
                        {
                            atws_l = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_atws_left_column]));
                        }
                        if (dt_library.Rows[j][wksp_atws_right_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_atws_right_column])) == true)
                        {
                            atws_r = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_atws_right_column]));
                        }
                        if (dt_library.Rows[j][wksp_tws_left_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_tws_left_column])) == true)
                        {
                            tws_l = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_tws_left_column]));
                        }
                        if (dt_library.Rows[j][wksp_tws_right_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_tws_right_column])) == true)
                        {
                            tws_r = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_tws_right_column]));
                        }
                        if (dt_library.Rows[j][wksp_perm_left_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_perm_left_column])) == true)
                        {
                            perm_l = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_perm_left_column]));
                        }
                        if (dt_library.Rows[j][wksp_perm_right_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_perm_right_column])) == true)
                        {
                            perm_r = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_perm_right_column]));
                        }


                        if (flip1 == true)
                        {
                            if (atws_l > 0 || atws_r > 0)
                            {
                                double t = atws_l;
                                atws_l = atws_r;
                                atws_r = t;
                            }
                            if (tws_l > 0 || tws_r > 0)
                            {
                                double t = tws_l;
                                tws_l = tws_r;
                                tws_r = t;
                            }
                            if (perm_l > 0 || perm_r > 0)
                            {
                                double t = perm_l;
                                perm_l = perm_r;
                                perm_r = t;
                            }
                        }
                        j = dt_library.Rows.Count;
                    }
                }
            }
        }

        private void build_dt_pcn()
        {


            dt_pcn = new System.Data.DataTable();

            dt_pcn.Columns.Add("p_perm_l", typeof(double));
            dt_pcn.Columns.Add("p_perm_r", typeof(double));
            dt_pcn.Columns.Add("p_tws_l", typeof(double));
            dt_pcn.Columns.Add("p_tws_r", typeof(double));
            dt_pcn.Columns.Add("p_atws_l", typeof(double));
            dt_pcn.Columns.Add("p_atws_r", typeof(double));

            dt_pcn.Columns.Add("perm_l", typeof(double));
            dt_pcn.Columns.Add("perm_r", typeof(double));
            dt_pcn.Columns.Add("tws_l", typeof(double));
            dt_pcn.Columns.Add("tws_r", typeof(double));
            dt_pcn.Columns.Add("atws_l", typeof(double));
            dt_pcn.Columns.Add("atws_r", typeof(double));

            dt_pcn.Columns.Add("n_perm_l", typeof(double));
            dt_pcn.Columns.Add("n_perm_r", typeof(double));
            dt_pcn.Columns.Add("n_tws_l", typeof(double));
            dt_pcn.Columns.Add("n_tws_r", typeof(double));
            dt_pcn.Columns.Add("n_atws_l", typeof(double));
            dt_pcn.Columns.Add("n_atws_r", typeof(double));

            double p_atws_l = 0;
            double p_tws_l = 0;
            double p_perm_l = 0;
            double p_atws_r = 0;
            double p_tws_r = 0;
            double p_perm_r = 0;

            if(dt_corridor==null)
            {
                //build dt corridor as default;
            }

            
            for (int i = 0; i < dt_corridor.Rows.Count; ++i)
            {


                if (dt_corridor.Rows[i][col_corridor_name] != DBNull.Value)
                {
                    double atws_l = 0;
                    double tws_l = 0;
                    double perm_l = 0;
                    double atws_r = 0;
                    double tws_r = 0;
                    double perm_r = 0;
                    string ws_name = Convert.ToString(dt_corridor.Rows[i][col_corridor_name]);
                    bool flipped = false;

                    if (dt_corridor.Rows[i][tws_side_column] != DBNull.Value && Convert.ToString(dt_corridor.Rows[i][tws_side_column]) == "flipped")
                    {
                        flipped = true;
                    }
                    set_corridor_offsets(ref perm_l, ref tws_l, ref atws_l, ref perm_r, ref tws_r, ref atws_r, ws_name, flipped);

                    if (i > 0)
                    {
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][12] = perm_l;
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][13] = perm_r;
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][14] = tws_l;
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][15] = tws_r;
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][16] = atws_l;
                        dt_pcn.Rows[dt_pcn.Rows.Count - 1][17] = atws_r;
                    }

                    dt_pcn.Rows.Add();

                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][0] = p_perm_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][1] = p_perm_r;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][2] = p_tws_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][3] = p_tws_r;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][4] = p_atws_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][5] = p_atws_r;

                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][6] = perm_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][7] = perm_r;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][8] = tws_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][9] = tws_r;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][10] = atws_l;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][11] = atws_r;

                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][12] = 0;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][13] = 0;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][14] = 0;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][15] = 0;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][16] = 0;
                    dt_pcn.Rows[dt_pcn.Rows.Count - 1][17] = 0;

                    p_atws_l = atws_l;
                    p_atws_r = atws_r;
                    p_tws_l = tws_l;
                    p_tws_r = tws_r;
                    p_perm_l = perm_l;
                    p_perm_r = perm_r;

                }
            }
        }


        public Polyline make_polyline_start_point_outside_corridor(Polyline bottom1, Polyline top1, Polyline feature1)
        {
            try
            {
                if (feature1.NumberOfVertices > 4)
                {
                    Polyline poly1 = new Polyline();
                    int idx = 0;
                    for (int i = 0; i < top1.NumberOfVertices; ++i)
                    {
                        poly1.AddVertexAt(idx, top1.GetPoint2dAt(i), top1.GetBulgeAt(i), 0, 0);
                        ++idx;
                    }
                    for (int i = bottom1.NumberOfVertices - 1; i >= 0; --i)
                    {
                        poly1.AddVertexAt(idx, bottom1.GetPoint2dAt(i), bottom1.GetBulgeAt(i), 0, 0);
                        ++idx;
                    }
                    poly1.Closed = true;

                    System.Data.DataTable dt1 = build_data_table_from_poly(feature1);

                    DBObjectCollection Poly_Colection = new DBObjectCollection();
                    Poly_Colection.Add(poly1);
                    DBObjectCollection Region_Colectionft = new DBObjectCollection();
                    Region_Colectionft = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection);

                    Region reg1 = Region_Colectionft[0] as Autodesk.AutoCAD.DatabaseServices.Region;

                    Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;
                    Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;

                    using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                    {
                        if (Brep_obj != null)
                        {
                            using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(feature1.StartPoint, out pc_start))
                            {
                                if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                {
                                    pc_start = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                }
                            }
                            using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent2 = Brep_obj.GetPointContainment(feature1.EndPoint, out pc_end))
                            {
                                if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                {
                                    pc_end = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                }
                            }

                        }
                    }

                    if (pc_start != Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside || pc_end != Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside)
                    {
                        Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;
                        Autodesk.AutoCAD.BoundaryRepresentation.PointContainment pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside;

                        bool run_brep = true;
                        int run_no = 0;

                        do
                        {
                            System.Data.DataRow row0 = dt1.Rows[0];
                            System.Data.DataRow row1 = dt1.NewRow();
                            row1.ItemArray = row0.ItemArray;
                            dt1.Rows[0].Delete();

                            dt1.Rows.Add();
                            for (int j = 0; j < dt1.Columns.Count; ++j)
                            {
                                dt1.Rows[dt1.Rows.Count - 1][j] = row1[j];
                            }

                            Point2d pt1 = (Point2d)dt1.Rows[0][0];
                            Point2d pt2 = (Point2d)dt1.Rows[dt1.Rows.Count - 1][0];

                            Point3d pt11 = new Point3d(pt1.X, pt1.Y, bottom1.Elevation);
                            Point3d pt22 = new Point3d(pt2.X, pt2.Y, bottom1.Elevation);

                            using (Autodesk.AutoCAD.BoundaryRepresentation.Brep Brep_obj = new Autodesk.AutoCAD.BoundaryRepresentation.Brep(reg1))
                            {
                                if (Brep_obj != null)
                                {
                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent1 = Brep_obj.GetPointContainment(pt11, out pc1))
                                    {
                                        if (ent1 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                        {
                                            pc1 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                        }
                                    }
                                    using (Autodesk.AutoCAD.BoundaryRepresentation.BrepEntity ent2 = Brep_obj.GetPointContainment(pt22, out pc2))
                                    {
                                        if (ent2 is Autodesk.AutoCAD.BoundaryRepresentation.Face)
                                        {
                                            pc2 = Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Inside;
                                        }
                                    }
                                }
                            }

                            if (pc1 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside && pc2 == Autodesk.AutoCAD.BoundaryRepresentation.PointContainment.Outside)
                            {
                                run_brep = false;
                            }
                            ++run_no;
                            if (run_no == 100) run_brep = false;
                        }
                        while (run_brep == true);

                        Polyline feature2 = new Polyline();

                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            feature2.AddVertexAt(i, (Point2d)dt1.Rows[i][0], 0, 0, 0);
                        }
                        return feature2;
                    }
                    else
                    {
                        return feature1;
                    }
                }
                else
                {
                    return feature1;
                }
            }
            catch (System.Exception)
            {

                return feature1;
            }
        }


        private Point3d build_crossover_feature(Polyline poly_cl, ref Polyline feature1, Point3d pt_pick, double stam)
        {
            //from point pick draft a line perp on cl then extend it other side to intersect lod3. from it calc p1 and p2. don't go closest

            Point3d pt1 = new Point3d(-1.234, -1.234, 0);

            int idx1 = 0;
            Polyline poly_perm_right = new Polyline();
            for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
            {
                if (dt_lod_right.Rows[i][6] != DBNull.Value)
                {
                    poly_perm_right.AddVertexAt(idx1, (Point2d)dt_lod_right.Rows[i][6], 0, 0, 0);
                    ++idx1;
                }
            }
            idx1 = 0;
            Polyline poly_perm_left = new Polyline();
            for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
            {
                if (dt_lod_left.Rows[i][6] != DBNull.Value)
                {
                    poly_perm_left.AddVertexAt(idx1, (Point2d)dt_lod_left.Rows[i][6], 0, 0, 0);
                    ++idx1;
                }
            }




            if (poly_perm_left.NumberOfVertices > 2 && poly_perm_left.NumberOfVertices > 2)
            {


                Point3d p1 = new Point3d();
                Point3d p2 = new Point3d();

                Polyline lod3_left = create_lod_construction_polylines(2, "LEFT");
                Polyline lod3_right = create_lod_construction_polylines(2, "RIGHT");


                Point3d pt_on_lod_left = lod3_left.GetClosestPointTo(pt_pick, Vector3d.ZAxis, false);
                double dist_scale1 = Math.Pow(Math.Pow(pt_pick.X - pt_on_lod_left.X, 2) + Math.Pow(pt_pick.Y - pt_on_lod_left.Y, 2), 0.5);

                Point3d pt_on_lod_right = lod3_right.GetClosestPointTo(pt_pick, Vector3d.ZAxis, false);
                double dist_scale2 = Math.Pow(Math.Pow(pt_pick.X - pt_on_lod_right.X, 2) + Math.Pow(pt_pick.Y - pt_on_lod_right.Y, 2), 0.5);

                Polyline poly_perm = poly_perm_left;
                Polyline lod3 = lod3_left;
                bool is_left = true;

                if (dist_scale2 < dist_scale1)
                {
                    poly_perm = poly_perm_right;
                    lod3 = lod3_right;
                    is_left = false;
                }

                for (int i = 0; i < dt_corridor.Rows.Count; ++i)
                {
                    if (is_left == false)
                    {
                        double sta_end = Convert.ToDouble(dt_corridor.Rows[i][tws_sta2_column]);

                        double param_m = Math.Round(poly_cl.GetParameterAtDistance(stam), 0);
                        double sta_pick = poly_cl.GetDistanceAtParameter(param_m);

                        if (Math.Round(sta_end, 0) == Math.Round(sta_pick, 0))
                        {
                            Point3d p_on_lod = lod3.GetClosestPointTo(pt_pick, Vector3d.ZAxis, false);

                            double paramp2 = Math.Round(lod3.GetParameterAtPoint(p_on_lod), 0);
                            double paramp1 = paramp2 - 1;

                            p1 = lod3.GetPointAtParameter(paramp1);
                            p2 = lod3.GetPointAtParameter(paramp2);
                            i = dt_corridor.Rows.Count;

                        }

                    }
                    else
                    {
                        double sta_start = Convert.ToDouble(dt_corridor.Rows[i][tws_sta1_column]);

                        double param_m = Math.Round(poly_cl.GetParameterAtDistance(stam), 0);
                        double sta_pick = poly_cl.GetDistanceAtParameter(param_m);

                        if (Math.Round(sta_start, 0) == Math.Round(sta_pick, 0))
                        {
                            Point3d p_on_lod = lod3.GetClosestPointTo(pt_pick, Vector3d.ZAxis, false);

                            double paramp2 = Math.Round(lod3.GetParameterAtPoint(p_on_lod), 0);
                            double paramp1 = paramp2 + 1;

                            p1 = lod3.GetPointAtParameter(paramp1);
                            p2 = lod3.GetPointAtParameter(paramp2);
                            i = dt_corridor.Rows.Count;

                        }

                    }

                }


                Polyline poly1 = new Polyline();
                poly1.AddVertexAt(0, new Point2d(p1.X, p1.Y), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(p2.X, p2.Y), 0, 0, 0);

                poly1.TransformBy(Matrix3d.Scaling((poly1.Length + 10 * dist_scale1) / poly1.Length, p1));

                Point3dCollection col_int1 = Functions.Intersect_on_both_operands(poly1, poly_perm);
                if (col_int1.Count == 1)
                {
                    Point3d pt_i1 = col_int1[0];

                    double d1 = Math.Pow(Math.Pow(p2.X - pt_i1.X, 2) + Math.Pow(p2.Y - pt_i1.Y, 2), 0.5);

                    Point3d pt_f2 = col_int1[0];

                    double param2 = Math.Round(poly_perm.GetParameterAtPoint(pt_f2), 0);
                    Point3d pt_f3 = poly_perm.GetPointAtParameter(param2);

                    Point3d pt_on_cl = poly_cl.GetClosestPointTo(pt_f3, Vector3d.ZAxis, false);

                    double param3 = Math.Round(poly_cl.GetParameterAtPoint(pt_on_cl), 0);
                    Point3d pt_f4 = poly_cl.GetPointAtParameter(param3);

                    Polyline poly2 = new Polyline();
                    poly2.AddVertexAt(0, new Point2d(pt_f3.X, pt_f3.Y), 0, 0, 0);
                    poly2.AddVertexAt(1, new Point2d(pt_f4.X, pt_f4.Y), 0, 0, 0);

                    poly2.TransformBy(Matrix3d.Scaling((poly2.Length + 1000) / poly2.Length, pt_f3));

                    Point3d pt_f5 = poly2.EndPoint;


                    feature1.AddVertexAt(0, new Point2d(p2.X, p2.Y), 0, 0, 0);
                    feature1.AddVertexAt(1, new Point2d(pt_f2.X, pt_f2.Y), 0, 0, 0);
                    feature1.AddVertexAt(2, new Point2d(pt_f3.X, pt_f3.Y), 0, 0, 0);
                    feature1.AddVertexAt(3, new Point2d(pt_f4.X, pt_f4.Y), 0, 0, 0);
                    feature1.AddVertexAt(4, new Point2d(pt_f5.X, pt_f5.Y), 0, 0, 0);




                    System.Data.DataTable dt_feat_side = build_data_table_from_poly(feature1);

                    Point3dCollection col_int_feature = Functions.Intersect_on_both_operands(poly_cl, feature1);
                    if (col_int_feature.Count == 1)
                    {
                        pt1 = col_int_feature[0];
                    }


                    string col1 = Convert.ToString(Math.Round(poly_cl.GetDistAtPoint(pt1), 2));
                    if (dt_sides == null)
                    {
                        dt_sides = new System.Data.DataTable();
                    }

                    if (dt_sides.Columns.Contains(col1) == false)
                    {
                        dt_sides.Columns.Add(col1, typeof(Point2d));
                        dt_sides.Columns.Add(col1 + "#", typeof(Point2d));
                    }
                    else
                    {
                        for (int i = 0; i < dt_feat_side.Rows.Count; ++i)
                        {
                            dt_sides.Rows[i][col1] = DBNull.Value;
                            dt_sides.Rows[0][col1 + "#"] = DBNull.Value;
                        }
                    }

                    for (int i = 0; i < dt_feat_side.Rows.Count; ++i)
                    {
                        if (dt_sides.Rows.Count == i)
                        {
                            dt_sides.Rows.Add();
                        }
                        dt_sides.Rows[i][col1] = dt_feat_side.Rows[i][0];
                    }
                    dt_sides.Rows[0][col1 + "#"] = new Point2d(pt1.X, pt1.Y);
                }
            }

            return pt1;
        }

    }

}

