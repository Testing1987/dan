using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class wksp_tool
    {
        public static Polyline create_lod_construction_polylines(int index1, string lr)
        {

            int indx = 0;
            Point2d pt_prev = new Point2d();
            Polyline lod1_left = null;
            Polyline lod1_right = null;
            Polyline lod2_left = null;
            Polyline lod2_right = null;
            Polyline lod3_left = null;
            Polyline lod3_right = null;
            Polyline lod4_left = null;
            Polyline lod4_right = null;


            if (index1 == 1 && lr == "LEFT")
            {
                if (dt_lod_left != null && dt_lod_left.Rows.Count > 1)
                {
                    lod1_left = new Polyline();
                    for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
                    {
                        if (dt_lod_left.Rows[i][0] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_left.Rows[i][0];
                            double bulge1 = 0;
                            if (dt_lod_left.Rows[i][1] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_left.Rows[i][1]);
                            }

                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod1_left.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod1_left;
            }


            else if (index1 == 1 && lr == "RIGHT")
            {
                if (dt_lod_right != null && dt_lod_right.Rows.Count > 1)
                {
                    indx = 0;
                    lod1_right = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
                    {
                        if (dt_lod_right.Rows[i][0] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_right.Rows[i][0];
                            double bulge1 = 0;
                            if (dt_lod_right.Rows[i][1] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_right.Rows[i][1]);
                            }

                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod1_right.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod1_right;
            }

            else if (index1 == 2 && lr == "LEFT")
            {
                if (dt_lod_left != null && dt_lod_left.Rows.Count > 1)
                {
                    indx = 0;
                    lod2_left = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
                    {
                        if (dt_lod_left.Rows[i][2] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_left.Rows[i][2];
                            double bulge1 = 0;
                            if (dt_lod_left.Rows[i][3] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_left.Rows[i][3]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod2_left.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod2_left;
            }

            else if (index1 == 2 && lr == "RIGHT")
            {
                if (dt_lod_right != null && dt_lod_right.Rows.Count > 1)
                {
                    indx = 0;
                    lod2_right = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
                    {
                        if (dt_lod_right.Rows[i][2] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_right.Rows[i][2];
                            double bulge1 = 0;
                            if (dt_lod_right.Rows[i][3] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_right.Rows[i][3]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod2_right.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod2_right;
            }

            else if (index1 == 3 && lr == "LEFT")
            {
                if (dt_lod_left != null && dt_lod_left.Rows.Count > 1)
                {
                    indx = 0;
                    lod3_left = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
                    {
                        if (dt_lod_left.Rows[i][4] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_left.Rows[i][4];
                            double bulge1 = 0;
                            if (dt_lod_left.Rows[i][5] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_left.Rows[i][5]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod3_left.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod3_left;
            }

            else if (index1 == 3 && lr == "RIGHT")
            {
                if (dt_lod_right != null && dt_lod_right.Rows.Count > 1)
                {
                    indx = 0;
                    lod3_right = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
                    {
                        if (dt_lod_right.Rows[i][4] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_right.Rows[i][4];
                            double bulge1 = 0;
                            if (dt_lod_right.Rows[i][5] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_right.Rows[i][5]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod3_right.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod3_right;
            }
            else if (index1 == 4 && lr == "LEFT")
            {
                if (dt_lod_left != null && dt_lod_left.Rows.Count > 1)
                {
                    indx = 0;
                    lod4_left = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
                    {
                        if (dt_lod_left.Rows[i][6] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_left.Rows[i][6];
                            double bulge1 = 0;
                            if (dt_lod_left.Rows[i][7] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_left.Rows[i][7]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod4_left.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod4_left;
            }
            else if (index1 == 4 && lr == "RIGHT")
            {
                if (dt_lod_right != null && dt_lod_right.Rows.Count > 1)
                {
                    indx = 0;
                    lod4_right = new Polyline();
                    pt_prev = new Point2d();
                    for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
                    {
                        if (dt_lod_right.Rows[i][6] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_lod_right.Rows[i][6];
                            double bulge1 = 0;
                            if (dt_lod_right.Rows[i][7] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod_right.Rows[i][7]);
                            }
                            double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                            if (d2d > 0.001)
                            {
                                lod4_right.AddVertexAt(indx, pt1, bulge1, 0, 0);
                                pt_prev = pt1;
                                ++indx;
                            }
                        }
                    }
                }
                return lod4_right;
            }
            return null;
        }

        public void draw_lod_polylines(Transaction Trans1, BlockTableRecord BTrecord)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            string dwg1 = ThisDrawing.Database.OriginalFileName;

            #region LOD polyline
            if ((dt_lod_left != null && dt_lod_left.Rows.Count > 0) || (dt_lod_right != null && dt_lod_right.Rows.Count > 0))
            {
                delete_existing_LOD();
            }


            int indx = 0;
            string lod_layer = comboBox_layer_lod.Text;
            Functions.Creaza_layer(lod_layer, 6, true);

            #region LOD left polyline
            if (dt_lod_left != null && dt_lod_left.Rows.Count > 0)
            {
                Polyline lod_left = new Polyline();
                indx = 0;
                Point2d pt_prev = new Point2d();
                for (int i = 0; i < dt_lod_left.Rows.Count; ++i)
                {
                    if (dt_lod_left.Rows[i][6] != DBNull.Value)
                    {
                        Point2d pt1 = (Point2d)dt_lod_left.Rows[i][6];
                        double bulge1 = 0;

                        if (dt_lod_left.Rows[i][7] != DBNull.Value)
                        {
                            bulge1 = Convert.ToDouble(dt_lod_left.Rows[i][7]);
                        }
                        double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                        if (d2d > 0.001)
                        {
                            lod_left.AddVertexAt(indx, pt1, bulge1, 0, 0);
                            pt_prev = pt1;
                            ++indx;
                        }
                    }
                }

                lod_left.ColorIndex = 256;
                lod_left.Layer = lod_layer;
                BTrecord.AppendEntity(lod_left);
                Trans1.AddNewlyCreatedDBObject(lod_left, true);

                dt_erase.Rows.Add();
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_dwg] = dwg1;
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_objid] = lod_left.ObjectId;
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_layer] = lod_layer;



                using (DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable)
                {
                    ObjectIdCollection col1 = new ObjectIdCollection();
                    col1.Add(lod_left.ObjectId);
                    DrawOrderTable1.MoveToBottom(col1);
                }


            }
            #endregion

            #region LOD right polyline
            if (dt_lod_right != null && dt_lod_right.Rows.Count > 0)
            {
                Polyline lod_right = new Polyline();
                indx = 0;
                Point2d pt_prev = new Point2d();
                for (int i = 0; i < dt_lod_right.Rows.Count; ++i)
                {
                    if (dt_lod_right.Rows[i][6] != DBNull.Value)
                    {
                        Point2d pt1 = (Point2d)dt_lod_right.Rows[i][6];
                        double bulge1 = 0;

                        if (dt_lod_right.Rows[i][7] != DBNull.Value)
                        {
                            bulge1 = Convert.ToDouble(dt_lod_right.Rows[i][7]);
                        }
                        double d2d = Math.Pow(Math.Pow(pt1.X - pt_prev.X, 2) + Math.Pow(pt1.Y - pt_prev.Y, 2), 0.5);

                        if (d2d > 0.001)
                        {
                            lod_right.AddVertexAt(indx, pt1, bulge1, 0, 0);
                            pt_prev = pt1;
                            ++indx;
                        }
                    }
                }

                lod_right.ColorIndex = 256;
                lod_right.Layer = lod_layer;
                BTrecord.AppendEntity(lod_right);
                Trans1.AddNewlyCreatedDBObject(lod_right, true);

                dt_erase.Rows.Add();
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_dwg] = dwg1;
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_objid] = lod_right.ObjectId;
                dt_erase.Rows[dt_erase.Rows.Count - 1][col_layer] = lod_layer;

                using (DrawOrderTable DrawOrderTable1 = Trans1.GetObject(BTrecord.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable)
                {
                    ObjectIdCollection col1 = new ObjectIdCollection();
                    col1.Add(lod_right.ObjectId);
                    DrawOrderTable1.MoveToBottom(col1);
                }
            }
            #endregion
            #endregion
        }

        private void delete_existing_LOD()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            string dwg1 = ThisDrawing.Database.OriginalFileName;

            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    if (dt_erase != null && dt_erase.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt_erase.Rows.Count; i++)
                        {
                            if (dt_erase.Rows[i][col_dwg] != DBNull.Value)
                            {
                                string dwg2 = Convert.ToString(dt_erase.Rows[i][col_dwg]);
                                if (dwg1 == dwg2)
                                {
                                    ObjectId id1 = ObjectId.Null;

                                    if (dt_erase.Rows[i][col_objid] != DBNull.Value)
                                    {
                                        id1 = (ObjectId)dt_erase.Rows[i][col_objid];
                                    }
                                    if (id1 != ObjectId.Null || id1.IsErased == false)
                                    {
                                        Curve curve1 = Trans1.GetObject(id1, OpenMode.ForRead) as Curve;
                                        if (curve1 != null)
                                        {
                                            if (curve1.Layer == comboBox_layer_lod.Text)
                                            {
                                                curve1.UpgradeOpen();
                                                curve1.Erase();
                                            }
                                        }
                                    }
                                }

                            }
                        }

                        Trans1.Commit();
                    }
                }
              
            }
        }


        public System.Data.DataTable build_lod3_and_lod4_columns(System.Data.DataTable dt_atws_sorted, System.Data.DataTable dt_lod)
        {
            if (dt_lod != null && dt_lod.Rows.Count > 0 && dt_atws_sorted != null && dt_atws_sorted.Rows.Count > 0)
            {
                dt_atws_sorted = Functions.Sort_data_table(dt_atws_sorted, atws_sta1_column);
                Point2dCollection col3 = new Point2dCollection();
                Point2dCollection col4 = new Point2dCollection();
                Polyline lod2 = new Polyline();
                int indx = 0;
                for (int i = 0; i < dt_lod.Rows.Count; ++i)
                {
                    if (dt_lod.Rows[i][2] != DBNull.Value)
                    {
                        double bulge1 = 0;
                        if (dt_lod.Rows[i][3] != DBNull.Value)
                        {
                            bulge1 = Convert.ToDouble(dt_lod.Rows[i][3]);
                        }
                        lod2.AddVertexAt(indx, (Point2d)dt_lod.Rows[i][2], bulge1, 0, 0);
                        ++indx;
                    }
                }




                double last_param2 = -1;

                bool exista_atws = false;
                for (int i = 0; i < dt_atws_sorted.Rows.Count; ++i)
                {
                    string handle1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_handle_column]);
                    string abutter1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_abutter_column]);

                    if (abutter1 == "ATWS")
                    {
                        exista_atws = true;
                    }

                    if (dt_atws_lod_manual != null && dt_atws_lod_manual.Columns.Contains(handle1) == true && abutter1 == "TWS")
                    {
                        int last_index = -1;
                        for (int n = 0; n < dt_atws_lod_manual.Rows.Count; ++n)
                        {
                            if (dt_atws_lod_manual.Rows[n][handle1] == DBNull.Value)
                            {
                                last_index = n - 1;
                                n = dt_atws_lod_manual.Rows.Count;
                            }
                            else
                            {
                                last_index = n;
                            }
                        }

                        if (dt_atws_lod_manual.Rows[0][handle1] != DBNull.Value && dt_atws_lod_manual.Rows[last_index][handle1] != DBNull.Value)
                        {
                            Point2d pt1 = (Point2d)dt_atws_lod_manual.Rows[0][handle1];
                            Point2d pt2 = (Point2d)dt_atws_lod_manual.Rows[last_index][handle1];

                            Point3d pt_on_poly1 = lod2.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                            Point3d pt_on_poly2 = lod2.GetClosestPointTo(new Point3d(pt2.X, pt2.Y, 0), Vector3d.ZAxis, false);
                            double param1 = lod2.GetParameterAtPoint(pt_on_poly1);
                            double param2 = lod2.GetParameterAtPoint(pt_on_poly2);


                            for (int k = 0; k < param1; ++k)
                            {
                                if (k > last_param2)
                                {
                                    col3.Add(lod2.GetPoint2dAt(k));
                                }
                            }

                            Point2d start1 = new Point2d(pt_on_poly1.X, pt_on_poly1.Y);
                            col3.Add(start1);
                            for (int j = 0; j < dt_atws_lod_manual.Rows.Count; ++j)
                            {
                                if (dt_atws_lod_manual.Rows[j][handle1] != DBNull.Value)
                                {
                                    col3.Add((Point2d)dt_atws_lod_manual.Rows[j][handle1]);
                                }

                            }
                            Point2d end1 = new Point2d(pt_on_poly2.X, pt_on_poly2.Y);
                            col3.Add(end1);
                            last_param2 = param2;
                        }

                    }

                }

                if (last_param2 != -1)
                {
                    for (int k = 0; k < lod2.NumberOfVertices; ++k)
                    {
                        if (k > last_param2)
                        {
                            col3.Add(lod2.GetPoint2dAt(k));
                        }
                    }
                }

                if (col3.Count > 0)
                {
                    for (int i = 0; i < dt_lod.Rows.Count; ++i)
                    {
                        dt_lod.Rows[i][4] = DBNull.Value;
                        dt_lod.Rows[i][5] = DBNull.Value;
                        dt_lod.Rows[i][6] = DBNull.Value;
                        dt_lod.Rows[i][7] = DBNull.Value;
                    }



                    for (int i = 0; i < col3.Count; ++i)
                    {
                        if (i == dt_lod.Rows.Count)
                        {
                            dt_lod.Rows.Add();
                        }
                        dt_lod.Rows[i][4] = col3[i];
                        dt_lod.Rows[i][5] = 0;
                        dt_lod.Rows[i][6] = col3[i];
                        dt_lod.Rows[i][7] = 0;
                    }

                    Polyline lod3 = new Polyline();
                    indx = 0;

                    for (int i = 0; i < dt_lod.Rows.Count; ++i)
                    {
                        if (dt_lod.Rows[i][4] != DBNull.Value)
                        {
                            double bulge1 = 0;
                            if (dt_lod.Rows[i][5] != DBNull.Value)
                            {
                                bulge1 = Convert.ToDouble(dt_lod.Rows[i][5]);
                            }
                            lod3.AddVertexAt(indx, (Point2d)dt_lod.Rows[i][4], bulge1, 0, 0);
                            ++indx;
                        }
                    }

                    if (exista_atws == true)
                    {
                        last_param2 = -1;

                        for (int i = 0; i < dt_atws_sorted.Rows.Count; ++i)
                        {
                            string handle1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_handle_column]);
                            string abutter1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_abutter_column]);


                            if (dt_atws_lod_manual != null && dt_atws_lod_manual.Columns.Contains(handle1) == true && abutter1 == "ATWS")
                            {
                                int last_index = -1;
                                for (int n = 0; n < dt_atws_lod_manual.Rows.Count; ++n)
                                {
                                    if (dt_atws_lod_manual.Rows[n][handle1] == DBNull.Value)
                                    {
                                        last_index = n - 1;
                                        n = dt_atws_lod_manual.Rows.Count;
                                    }
                                    else
                                    {
                                        last_index = n;
                                    }
                                }

                                Point2d pt1 = (Point2d)dt_atws_lod_manual.Rows[0][handle1];
                                Point2d pt2 = (Point2d)dt_atws_lod_manual.Rows[last_index][handle1];

                                Point3d pt_on_poly1 = lod3.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                                Point3d pt_on_poly2 = lod3.GetClosestPointTo(new Point3d(pt2.X, pt2.Y, 0), Vector3d.ZAxis, false);
                                double param1 = lod3.GetParameterAtPoint(pt_on_poly1);
                                double param2 = lod3.GetParameterAtPoint(pt_on_poly2);


                                for (int k = 0; k < param1; ++k)
                                {
                                    if (k > last_param2)
                                    {
                                        col4.Add(lod3.GetPoint2dAt(k));
                                    }
                                }

                                Point2d start1 = new Point2d(pt_on_poly1.X, pt_on_poly1.Y);
                                col4.Add(start1);
                                for (int j = 0; j < dt_atws_lod_manual.Rows.Count; ++j)
                                {
                                    if (dt_atws_lod_manual.Rows[j][handle1] != DBNull.Value)
                                    {
                                        col4.Add((Point2d)dt_atws_lod_manual.Rows[j][handle1]);
                                    }

                                }
                                Point2d end1 = new Point2d(pt_on_poly2.X, pt_on_poly2.Y);
                                col4.Add(end1);
                                last_param2 = param2;
                            }

                        }

                        if (last_param2 != -1)
                        {
                            for (int k = 0; k < lod3.NumberOfVertices; ++k)
                            {
                                if (k > last_param2)
                                {
                                    col4.Add(lod3.GetPoint2dAt(k));
                                }
                            }
                        }

                        if (col4.Count > 0)
                        {
                            for (int i = 0; i < dt_lod.Rows.Count; ++i)
                            {
                                dt_lod.Rows[i][6] = DBNull.Value;
                                dt_lod.Rows[i][7] = DBNull.Value;
                            }

                            for (int i = 0; i < col4.Count; ++i)
                            {
                                if (i == dt_lod.Rows.Count)
                                {
                                    dt_lod.Rows.Add();
                                }

                                dt_lod.Rows[i][6] = col4[i];
                                dt_lod.Rows[i][7] = 0;
                            }
                        }


                    }


                }

            }

            return dt_lod;
        }

        public System.Data.DataTable build_lod_datatable(System.Data.DataTable dt_atws_sorted, System.Data.DataTable dt_lod, Polyline lod1, Polyline lod2)
        {
            if (dt_lod != null && dt_lod.Rows.Count > 1)
            {


                for (int i = 0; i < dt_lod.Rows.Count; ++i)
                {
                    dt_lod.Rows[i][0] = DBNull.Value;
                    dt_lod.Rows[i][1] = DBNull.Value;
                    dt_lod.Rows[i][2] = DBNull.Value;
                    dt_lod.Rows[i][3] = DBNull.Value;
                    dt_lod.Rows[i][4] = DBNull.Value;
                    dt_lod.Rows[i][5] = DBNull.Value;
                    dt_lod.Rows[i][6] = DBNull.Value;
                    dt_lod.Rows[i][7] = DBNull.Value;
                }




                for (int i = 0; i < lod1.NumberOfVertices; ++i)
                {
                    if (i == dt_lod.Rows.Count)
                    {
                        dt_lod.Rows.Add();
                    }
                    dt_lod.Rows[i][0] = lod1.GetPoint2dAt(i);
                    dt_lod.Rows[i][1] = 0;
                }

                for (int i = 0; i < lod2.NumberOfVertices; ++i)
                {
                    if (i == dt_lod.Rows.Count)
                    {
                        dt_lod.Rows.Add();
                    }
                    dt_lod.Rows[i][2] = lod2.GetPoint2dAt(i);
                    dt_lod.Rows[i][3] = 0;
                }

                if (dt_atws_sorted != null && dt_atws_sorted.Rows.Count > 0)
                {
                    Polyline lod3 = new Polyline();
                    Polyline lod4 = new Polyline();

                    for (int i = 0; i < lod2.NumberOfVertices; ++i)
                    {
                        lod3.AddVertexAt(i, lod2.GetPoint2dAt(i), lod2.GetBulgeAt(i), 0, 0);
                    }


                    dt_atws_sorted = Functions.Sort_data_table(dt_atws_sorted, atws_sta1_column);
                    for (int i = 0; i < dt_atws_sorted.Rows.Count; ++i)
                    {
                        string handle1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_handle_column]);
                        string abutter1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_abutter_column]);

                        if (dt_atws_lod_manual != null && dt_atws_lod_manual.Columns.Contains(handle1) == true && abutter1 == "TWS")
                        {
                            Polyline poly_temp = new Polyline();
                            int index1 = 0;

                            for (int j = 0; j < dt_atws_lod_manual.Rows.Count; ++j)
                            {
                                if (dt_atws_lod_manual.Rows[j][handle1] != DBNull.Value)
                                {
                                    poly_temp.AddVertexAt(index1, (Point2d)dt_atws_lod_manual.Rows[j][handle1], 0, 0, 0);
                                    ++index1;
                                }
                            }

                            Point3d pt_on_poly1 = lod3.GetClosestPointTo(poly_temp.StartPoint, Vector3d.ZAxis, false);
                            Point3d pt_on_poly2 = lod3.GetClosestPointTo(poly_temp.EndPoint, Vector3d.ZAxis, false);
                            double param1 = lod3.GetParameterAtPoint(pt_on_poly1);
                            double param2 = lod3.GetParameterAtPoint(pt_on_poly2);

                            Polyline start1 = get_part_of_poly(lod3, 0, param1);
                            Polyline end1 = get_part_of_poly(lod3, param2, lod3.EndParam);

                            lod3 = new Polyline();
                            index1 = 0;

                            if (param1 > 0)
                            {

                                for (int j = 0; j < start1.NumberOfVertices; ++j)
                                {
                                    lod3.AddVertexAt(index1, start1.GetPoint2dAt(j), start1.GetBulgeAt(j), 0, 0);
                                    ++index1;
                                }
                            }
                            for (int j = 0; j < poly_temp.NumberOfVertices; ++j)
                            {
                                lod3.AddVertexAt(index1, poly_temp.GetPoint2dAt(j), poly_temp.GetBulgeAt(j), 0, 0);
                                ++index1;
                            }

                            if (param2 < lod3.EndParam)
                            {

                                for (int j = 0; j < end1.NumberOfVertices; ++j)
                                {
                                    lod3.AddVertexAt(index1, end1.GetPoint2dAt(j), end1.GetBulgeAt(j), 0, 0);
                                    ++index1;
                                }
                            }
                        }

                    }

                    for (int i = 0; i < lod3.NumberOfVertices; ++i)
                    {
                        if (i == dt_lod.Rows.Count)
                        {
                            dt_lod.Rows.Add();
                        }
                        dt_lod.Rows[i][4] = lod3.GetPoint2dAt(i);
                        dt_lod.Rows[i][5] = 0;
                    }

                    for (int i = 0; i < lod3.NumberOfVertices; ++i)
                    {
                        lod4.AddVertexAt(i, lod3.GetPoint2dAt(i), lod3.GetBulgeAt(i), 0, 0);
                    }

                    for (int i = 0; i < dt_atws_sorted.Rows.Count; ++i)
                    {
                        string handle1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_handle_column]);
                        string abutter1 = Convert.ToString(dt_atws_sorted.Rows[i][atws_abutter_column]);

                        if (dt_atws_lod_manual != null && dt_atws_lod_manual.Columns.Contains(handle1) == true && abutter1 == "ATWS")
                        {
                            Polyline poly_temp = new Polyline();
                            int index1 = 0;

                            for (int j = 0; j < dt_atws_lod_manual.Rows.Count; ++j)
                            {
                                if (dt_atws_lod_manual.Rows[j][handle1] != DBNull.Value)
                                {
                                    poly_temp.AddVertexAt(index1, (Point2d)dt_atws_lod_manual.Rows[j][handle1], 0, 0, 0);
                                    ++index1;
                                }
                            }

                            Point3d pt_on_poly1 = lod4.GetClosestPointTo(poly_temp.StartPoint, Vector3d.ZAxis, false);
                            Point3d pt_on_poly2 = lod4.GetClosestPointTo(poly_temp.EndPoint, Vector3d.ZAxis, false);
                            double param1 = lod4.GetParameterAtPoint(pt_on_poly1);
                            double param2 = lod4.GetParameterAtPoint(pt_on_poly2);

                            Polyline start1 = get_part_of_poly(lod4, 0, param1);
                            Polyline end1 = get_part_of_poly(lod4, param2, lod4.EndParam);

                            lod4 = new Polyline();
                            index1 = 0;

                            if (param1 > 0)
                            {
                                for (int j = 0; j < start1.NumberOfVertices; ++j)
                                {
                                    lod4.AddVertexAt(index1, start1.GetPoint2dAt(j), start1.GetBulgeAt(j), 0, 0);
                                    ++index1;
                                }
                            }
                            for (int j = 0; j < poly_temp.NumberOfVertices; ++j)
                            {
                                lod4.AddVertexAt(index1, poly_temp.GetPoint2dAt(j), poly_temp.GetBulgeAt(j), 0, 0);
                                ++index1;
                            }

                            if (param2 < lod4.EndParam)
                            {
                                for (int j = 0; j < end1.NumberOfVertices; ++j)
                                {
                                    lod4.AddVertexAt(index1, end1.GetPoint2dAt(j), end1.GetBulgeAt(j), 0, 0);
                                    ++index1;
                                }
                            }
                        }
                    }
                    for (int i = 0; i < lod4.NumberOfVertices; ++i)
                    {
                        if (i == dt_lod.Rows.Count)
                        {
                            dt_lod.Rows.Add();
                        }
                        dt_lod.Rows[i][6] = lod4.GetPoint2dAt(i);
                        dt_lod.Rows[i][7] = 0;
                    }
                }
                else
                {
                    for (int i = dt_lod.Rows.Count - 1; i >= 0; --i)
                    {
                        if (dt_lod.Rows[i][2] != DBNull.Value)
                        {
                            dt_lod.Rows[i][4] = dt_lod.Rows[i][2];
                            dt_lod.Rows[i][5] = dt_lod.Rows[i][3];
                            dt_lod.Rows[i][6] = dt_lod.Rows[i][2];
                            dt_lod.Rows[i][7] = dt_lod.Rows[i][3];
                        }
                        else
                        {
                            dt_lod.Rows[i].Delete();
                        }
                    }
                }
            }
            return dt_lod;
        }


        private void calculate_LOD()
        {
            if (dt_cl == null || dt_cl.Rows.Count < 2 || dt_library == null || dt_library.Rows.Count == 0 || dt_corridor == null || dt_corridor.Rows.Count == 0)
            {
                return;
            }



            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (ThisDrawing == null)
            {
                set_enable_true();
                return;
            }

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


                        #region poly_dtcl
                        Polyline poly_dtcl = new Polyline();
                        for (int k = 0; k < dt_cl.Rows.Count; ++k)
                        {
                            if (dt_cl.Rows[k][0] == DBNull.Value || dt_cl.Rows[k][1] == DBNull.Value || Functions.IsNumeric(Convert.ToString(dt_cl.Rows[k][1])) == false)
                            {

                                set_enable_true();
                                return;
                            }
                            poly_dtcl.AddVertexAt(k, (Point2d)dt_cl.Rows[k][0], 0, 0, 0);
                        }
                        poly_dtcl.Elevation = 0;
                        #endregion

                        System.Data.DataTable dt_perm_left = new System.Data.DataTable();
                        System.Data.DataTable dt_perm_right = new System.Data.DataTable();
                        dt_perm_left.Columns.Add("pt", typeof(Point3d));
                        dt_perm_left.Columns.Add("idx", typeof(int));
                        dt_perm_right = dt_perm_left.Clone();

                        List<Point3dCollection> lista_tws_l = new List<Point3dCollection>();
                        List<Point3dCollection> lista_perm_for_tws_l = new List<Point3dCollection>();


                        List<Point3dCollection> lista_tws_r = new List<Point3dCollection>();
                        List<Point3dCollection> lista_perm_for_tws_r = new List<Point3dCollection>();


                        dt_lod_left = new System.Data.DataTable();

                        dt_lod_left.Columns.Add("pt1", typeof(Point2d));
                        dt_lod_left.Columns.Add("bulge1", typeof(double));

                        dt_lod_left.Columns.Add("pt2", typeof(Point2d));
                        dt_lod_left.Columns.Add("bulge2", typeof(double));

                        dt_lod_left.Columns.Add("pt3", typeof(Point2d));
                        dt_lod_left.Columns.Add("bulge3", typeof(double));

                        dt_lod_left.Columns.Add("pt4", typeof(Point2d));
                        dt_lod_left.Columns.Add("bulge4", typeof(double));

                        dt_lod_right = new System.Data.DataTable();

                        dt_lod_right.Columns.Add("pt1", typeof(Point2d));
                        dt_lod_right.Columns.Add("bulge1", typeof(double));

                        dt_lod_right.Columns.Add("pt2", typeof(Point2d));
                        dt_lod_right.Columns.Add("bulge2", typeof(double));

                        dt_lod_right.Columns.Add("pt3", typeof(Point2d));
                        dt_lod_right.Columns.Add("bulge3", typeof(double));

                        dt_lod_right.Columns.Add("pt4", typeof(Point2d));
                        dt_lod_right.Columns.Add("bulge4", typeof(double));

                        System.Data.DataTable dt_temp_lod_atws = new System.Data.DataTable();

                        for (int i = 0; i < dt_corridor.Rows.Count; ++i)
                        {
                            if (dt_corridor.Rows[i][tws_sta1_column] == DBNull.Value || Functions.IsNumeric(Convert.ToString(dt_corridor.Rows[i][tws_sta1_column])) == false ||
                                dt_corridor.Rows[i][tws_sta2_column] == DBNull.Value || Functions.IsNumeric(Convert.ToString(dt_corridor.Rows[i][tws_sta2_column])) == false)
                            {
                                MessageBox.Show("station start/end was not specified correctly");
                                set_enable_true();
                                return;
                            }

                            bool atws_modified = false;

                            if (dt_corridor.Rows[i][tws_modified_column] != DBNull.Value)
                            {
                                atws_modified = Convert.ToBoolean(dt_corridor.Rows[i][tws_modified_column]);
                            }

                            double sta1 = Math.Round(Convert.ToDouble(dt_corridor.Rows[i][tws_sta1_column]), 2);
                            double sta2 = Math.Round(Convert.ToDouble(dt_corridor.Rows[i][tws_sta2_column]), 2);
                            if (dt_corridor.Rows[i][col_corridor_name] != DBNull.Value)
                            {
                                double atws_width_left = 0;
                                double tws_l = 0;
                                double perm_l = 0;
                                double atws_width_right = 0;
                                double tws_r = 0;
                                double perm_r = 0;
                                string ws_name = Convert.ToString(dt_corridor.Rows[i][col_corridor_name]);

                                #region load parameters for permanent tws and atws

                                for (int j = 0; j < dt_library.Rows.Count; ++j)
                                {
                                    if (dt_library.Rows[j][lib_name_column] != DBNull.Value)
                                    {
                                        string nume1 = Convert.ToString(dt_library.Rows[j][lib_name_column]);
                                        if (nume1 == ws_name)
                                        {
                                            if (dt_library.Rows[j][wksp_atws_left_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_atws_left_column])) == true)
                                            {
                                                atws_width_left = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_atws_left_column]));
                                            }
                                            if (dt_library.Rows[j][wksp_atws_right_column] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_library.Rows[j][wksp_atws_right_column])) == true)
                                            {
                                                atws_width_right = Math.Abs(Convert.ToDouble(dt_library.Rows[j][wksp_atws_right_column]));
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

                                            if (checkBox_flip_workspace.Checked == true)
                                            {
                                                if (atws_width_left > 0 || atws_width_right > 0)
                                                {
                                                    double t = atws_width_left;
                                                    atws_width_left = atws_width_right;
                                                    atws_width_right = t;
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
                                #endregion

                                #region build temp polyline                                
                                if (sta1 > poly_dtcl.Length || sta2 > poly_dtcl.Length)
                                {

                                    set_enable_true();
                                    return;
                                }
                                double param1 = poly_dtcl.GetParameterAtDistance(sta1);
                                double param2 = poly_dtcl.GetParameterAtDistance(sta2);
                                Point3d pt1 = poly_dtcl.GetPointAtParameter(param1);
                                Point3d pt2 = poly_dtcl.GetPointAtParameter(param2);
                                #endregion


                                if (perm_l > 0 && perm_r > 0)
                                {


                                    Polyline poly_right_perm = get_trimmed_offset(poly_dtcl, sta1, sta2, perm_r);
                                    Polyline poly_left_perm = get_trimmed_offset(poly_dtcl, sta1, sta2, -perm_l);
                                    Polyline poly_left_tws = null;
                                    Polyline poly_right_tws = null;
                                    Polyline poly_left_atws = null;
                                    Polyline poly_right_atws = null;

                                    #region tws left
                                    if (tws_l > 0)
                                    {
                                        poly_left_tws = get_trimmed_offset(poly_dtcl, sta1, sta2, -perm_l - tws_l);

                                    }
                                    #endregion

                                    #region tws right
                                    if (tws_r > 0)
                                    {

                                        poly_right_tws = get_trimmed_offset(poly_dtcl, sta1, sta2, perm_r + tws_r);

                                    }
                                    #endregion

                                    #region atws left
                                    if (atws_width_left > 0)
                                    {

                                        poly_left_atws = get_trimmed_offset(poly_dtcl, sta1, sta2, -perm_l - tws_l - atws_width_left);

                                    }
                                    #endregion

                                    #region atws right
                                    if (atws_width_right > 0)
                                    {

                                        poly_right_atws = get_trimmed_offset(poly_dtcl, sta1, sta2, perm_r + tws_r + atws_width_right);


                                    }
                                    #endregion



                                    #region dt_LOD_left from corridor
                                    if (tws_l == 0 && atws_width_left == 0)
                                    {
                                        for (int k = 0; k < poly_left_perm.NumberOfVertices; ++k)
                                        {
                                            dt_lod_left.Rows.Add();
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][0] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][1] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][2] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][3] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][4] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][5] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][6] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][7] = 0;
                                        }
                                    }
                                    else if (tws_l > 0 && atws_width_left == 0)
                                    {
                                        for (int k = 0; k < poly_left_tws.NumberOfVertices; ++k)
                                        {
                                            dt_lod_left.Rows.Add();
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][0] = poly_left_tws.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][1] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][2] = poly_left_tws.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][3] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][4] = poly_left_tws.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][5] = 0;
                                        }


                                        int index_dt = dt_lod_left.Rows.Count;
                                        for (int k = 0; k < poly_left_perm.NumberOfVertices; ++k)
                                        {
                                            if (index_dt == dt_lod_left.Rows.Count)
                                            {
                                                dt_lod_left.Rows.Add();
                                            }
                                            dt_lod_left.Rows[index_dt][6] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[index_dt][7] = 0;
                                            ++index_dt;
                                        }


                                    }
                                    else if (atws_width_left > 0)
                                    {
                                        int index_dt = dt_lod_left.Rows.Count;
                                        for (int k = 0; k < poly_left_atws.NumberOfVertices; ++k)
                                        {
                                            dt_lod_left.Rows.Add();
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][2] = poly_left_atws.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][3] = 0;
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][4] = poly_left_atws.GetPoint2dAt(k);
                                            dt_lod_left.Rows[dt_lod_left.Rows.Count - 1][5] = 0;
                                        }

                                        if (tws_l == 0)
                                        {
                                            for (int k = 0; k < poly_left_perm.NumberOfVertices; ++k)
                                            {
                                                if (index_dt == dt_lod_left.Rows.Count)
                                                {
                                                    dt_lod_left.Rows.Add();
                                                }
                                                dt_lod_left.Rows[index_dt][0] = poly_left_perm.GetPoint2dAt(k);
                                                dt_lod_left.Rows[index_dt][1] = 0;
                                                ++index_dt;
                                            }
                                        }
                                        else
                                        {
                                            for (int k = 0; k < poly_left_tws.NumberOfVertices; ++k)
                                            {
                                                if (index_dt == dt_lod_left.Rows.Count)
                                                {
                                                    dt_lod_left.Rows.Add();
                                                }
                                                dt_lod_left.Rows[index_dt][0] = poly_left_tws.GetPoint2dAt(k);
                                                dt_lod_left.Rows[index_dt][1] = 0;
                                                ++index_dt;
                                            }
                                        }

                                        for (int k = 0; k < poly_left_perm.NumberOfVertices; ++k)
                                        {
                                            if (index_dt == dt_lod_left.Rows.Count)
                                            {
                                                dt_lod_left.Rows.Add();
                                            }
                                            dt_lod_left.Rows[index_dt][6] = poly_left_perm.GetPoint2dAt(k);
                                            dt_lod_left.Rows[index_dt][7] = 0;
                                            ++index_dt;
                                        }

                                    }



                                    #endregion

                                    #region dt_LOD_right from corridor
                                    if (tws_r == 0 && atws_width_right == 0)
                                    {
                                        for (int k = 0; k < poly_right_perm.NumberOfVertices; ++k)
                                        {
                                            dt_lod_right.Rows.Add();
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][0] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][1] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][2] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][3] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][4] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][5] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][6] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][7] = 0;
                                        }
                                    }
                                    else if (tws_r > 0 && atws_width_right == 0)
                                    {
                                        for (int k = 0; k < poly_right_tws.NumberOfVertices; ++k)
                                        {
                                            dt_lod_right.Rows.Add();
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][0] = poly_right_tws.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][1] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][2] = poly_right_tws.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][3] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][4] = poly_right_tws.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][5] = 0;
                                        }

                                        int index_dt = dt_lod_right.Rows.Count;
                                        for (int k = 0; k < poly_right_perm.NumberOfVertices; ++k)
                                        {
                                            if (index_dt == dt_lod_right.Rows.Count)
                                            {
                                                dt_lod_right.Rows.Add();
                                            }
                                            dt_lod_right.Rows[index_dt][6] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[index_dt][7] = 0;
                                            ++index_dt;
                                        }

                                    }
                                    else if (atws_width_right > 0)
                                    {
                                        int index_dt = dt_lod_right.Rows.Count;
                                        for (int k = 0; k < poly_right_atws.NumberOfVertices; ++k)
                                        {
                                            dt_lod_right.Rows.Add();
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][2] = poly_right_atws.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][3] = 0;
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][4] = poly_right_atws.GetPoint2dAt(k);
                                            dt_lod_right.Rows[dt_lod_right.Rows.Count - 1][5] = 0;
                                        }

                                        if (tws_r == 0)
                                        {
                                            for (int k = 0; k < poly_right_perm.NumberOfVertices; ++k)
                                            {
                                                if (index_dt == dt_lod_right.Rows.Count)
                                                {
                                                    dt_lod_right.Rows.Add();
                                                }
                                                dt_lod_right.Rows[index_dt][0] = poly_right_perm.GetPoint2dAt(k);
                                                dt_lod_right.Rows[index_dt][1] = 0;
                                                ++index_dt;
                                            }
                                        }
                                        else
                                        {
                                            for (int k = 0; k < poly_right_tws.NumberOfVertices; ++k)
                                            {
                                                if (index_dt == dt_lod_right.Rows.Count)
                                                {
                                                    dt_lod_right.Rows.Add();
                                                }
                                                dt_lod_right.Rows[index_dt][0] = poly_right_tws.GetPoint2dAt(k);
                                                dt_lod_right.Rows[index_dt][1] = 0;
                                                ++index_dt;
                                            }
                                        }

                                        for (int k = 0; k < poly_right_perm.NumberOfVertices; ++k)
                                        {
                                            if (index_dt == dt_lod_right.Rows.Count)
                                            {
                                                dt_lod_right.Rows.Add();
                                            }
                                            dt_lod_right.Rows[index_dt][6] = poly_right_perm.GetPoint2dAt(k);
                                            dt_lod_right.Rows[index_dt][7] = 0;
                                            ++index_dt;
                                        }

                                    }

                                    #endregion
                                }
                            }
                        }

                        #region LOD database
                        if (dt_atws != null && dt_atws.Rows.Count > 0)
                        {
                            System.Data.DataTable dt_sorted_right = dt_atws.Clone();
                            for (int i = 0; i < dt_atws.Rows.Count; ++i)
                            {
                                if (Convert.ToString(dt_atws.Rows[i][atws_working_side_column]) == "RIGHT")
                                {
                                    System.Data.DataRow row1 = dt_sorted_right.NewRow();
                                    row1.ItemArray = dt_atws.Rows[i].ItemArray;
                                    dt_sorted_right.Rows.InsertAt(row1, dt_sorted_right.Rows.Count);
                                }
                            }
                            dt_lod_right = build_lod3_and_lod4_columns(dt_sorted_right, dt_lod_right);


                            System.Data.DataTable dt_sorted_left = dt_atws.Clone();
                            for (int i = 0; i < dt_atws.Rows.Count; ++i)
                            {
                                if (Convert.ToString(dt_atws.Rows[i][atws_working_side_column]) == "LEFT")
                                {
                                    System.Data.DataRow row1 = dt_sorted_left.NewRow();
                                    row1.ItemArray = dt_atws.Rows[i].ItemArray;
                                    dt_sorted_left.Rows.InsertAt(row1, dt_sorted_left.Rows.Count);

                                }
                            }
                            dt_lod_left = build_lod3_and_lod4_columns(dt_sorted_left, dt_lod_left);
                        }
                        #endregion

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
