using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class class_commands
    {


        [CommandMethod("ws1")]
        public void code_test_offset()
        {
            if (ZZCommand_class.isSECURE() == false) return;


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


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        System.Data.DataTable dt1 = new System.Data.DataTable();

                        dt1.Columns.Add("perm_right", typeof(double));
                        dt1.Columns.Add("perm_left", typeof(double));

                        dt1.Columns.Add("tws_right", typeof(double));
                        dt1.Columns.Add("tws_left", typeof(double));


                        dt1.Rows.Add();
                        dt1.Rows[0][0] = 10;
                        dt1.Rows[0][1] = 30;
                        dt1.Rows[0][2] = 5;
                        dt1.Rows[0][3] = 5;

                        dt1.Rows.Add();
                        dt1.Rows[1][0] = 25;
                        dt1.Rows[1][1] = 35;
                        dt1.Rows[1][2] = 10;
                        dt1.Rows[1][3] = 15;


                        Polyline poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;

                        if (poly1 != null && dt1.Rows[0][0] != DBNull.Value && dt1.Rows[0][1] != DBNull.Value)
                        {
                            DBObjectCollection col_offset_right = poly1.GetOffsetCurves(Convert.ToDouble(dt1.Rows[0][0]));
                            DBObjectCollection col_offset_left = poly1.GetOffsetCurves(-Convert.ToDouble(dt1.Rows[0][1]));
                            Polyline poly_right = col_offset_right[0] as Polyline;
                            Polyline poly_left = col_offset_left[0] as Polyline;



                            if (poly_left != null && poly_right != null)
                            {

                                Point3dCollection col1 = new Point3dCollection();

                                for (int i = 0; i < poly_left.NumberOfVertices; ++i)
                                {
                                    col1.Add(poly_left.GetPointAtParameter(i));
                                }

                                for (int j = poly_right.NumberOfVertices - 1; j >= 0; --j)
                                {
                                    col1.Add(poly_right.GetPointAtParameter(j));
                                }

                                if (col1.Count > 1)
                                {
                                    Polyline poly_perm = new Polyline();
                                    for (int i = 0; i < col1.Count; ++i)
                                    {
                                        poly_perm.AddVertexAt(i, new Point2d(col1[i].X, col1[i].Y), 0, 0, 0);
                                    }
                                    poly_perm.Closed = true;
                                    poly_perm.ColorIndex = 1;
                                    BTrecord.AppendEntity(poly_perm);
                                    Trans1.AddNewlyCreatedDBObject(poly_perm, true);

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



        }

        [CommandMethod("tst123")]
        public void code_test_edit_perm()
        {
            if (ZZCommand_class.isSECURE() == false) return;


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("perm_right", typeof(double));
            dt1.Columns.Add("perm_left", typeof(double));

            dt1.Columns.Add("tws_right", typeof(double));
            dt1.Columns.Add("tws_left", typeof(double));


            dt1.Rows.Add();
            dt1.Rows[0][0] = 10;
            dt1.Rows[0][1] = 30;
            dt1.Rows[0][2] = 5;
            dt1.Rows[0][3] = 5;

            dt1.Rows.Add();
            dt1.Rows[1][0] = 25;
            dt1.Rows[1][1] = 35;
            dt1.Rows[1][2] = 10;
            dt1.Rows[1][3] = 15;

            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_cl;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_cl;
                        Prompt_cl = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_cl.SetRejectMessage("\nSelect a polyline!");
                        Prompt_cl.AllowNone = true;
                        Prompt_cl.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_cl = ThisDrawing.Editor.GetEntity(Prompt_cl);

                        if (Rezultat_cl.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_perm;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_perm;
                        Prompt_perm = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the permanent easement:");
                        Prompt_perm.SetRejectMessage("\nSelect a polyline!");
                        Prompt_perm.AllowNone = true;
                        Prompt_perm.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_perm = ThisDrawing.Editor.GetEntity(Prompt_perm);

                        if (Rezultat_perm.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the switching point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        Polyline poly_perm = Trans1.GetObject(Rezultat_perm.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline poly_cl = Trans1.GetObject(Rezultat_cl.ObjectId, OpenMode.ForRead) as Polyline;

                        if (poly_perm != null && poly_cl != null && dt1 != null && dt1.Rows.Count > 0 && dt1.Rows[1][0] != DBNull.Value && dt1.Rows[1][1] != DBNull.Value)
                        {

                            double sta_start_switch = poly_cl.GetDistAtPoint(poly_cl.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false));

                            DBObjectCollection col_offset_right = poly_cl.GetOffsetCurves(Convert.ToDouble(dt1.Rows[1][0]));
                            DBObjectCollection col_offset_left = poly_cl.GetOffsetCurves(-Convert.ToDouble(dt1.Rows[1][1]));
                            Polyline poly_right = col_offset_right[0] as Polyline;
                            Polyline poly_left = col_offset_left[0] as Polyline;

                            Point3dCollection col1 = new Point3dCollection();

                            Point3d pt1 = poly_cl.GetPointAtDist(sta_start_switch);
                            Point3d pt2L = poly_left.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                            Point3d pt2R = poly_right.GetClosestPointTo(pt1, Vector3d.ZAxis, false);



                            Xline xline1 = new Xline();
                            xline1.BasePoint = pt1;
                            xline1.SecondPoint = pt2L;

                            Point3dCollection colint_left_new_perm = Functions.Intersect_on_both_operands(xline1, poly_left);
                            Point3dCollection colint_left_old_perm = Functions.Intersect_on_both_operands(xline1, poly_perm);

                            Xline xline2 = new Xline();
                            xline2.BasePoint = pt1;
                            xline2.SecondPoint = pt2R;

                            Point3dCollection colint_right_new_perm = Functions.Intersect_on_both_operands(xline2, poly_right);
                            Point3dCollection colint_right_old_perm = Functions.Intersect_on_both_operands(xline2, poly_perm);


                            double sta1L = poly_left.GetDistAtPoint(colint_left_new_perm[0]);
                            double sta1R = poly_right.GetDistAtPoint(colint_right_new_perm[0]);

                            if (poly_left != null && poly_right != null && colint_left_new_perm.Count > 0 && colint_left_old_perm.Count > 0)
                            {


                                Point3d pt_int_new_left = colint_left_new_perm[0];
                                Point3d pt_int_old_left = new Point3d();
                                for (int i = 0; i < colint_left_old_perm.Count; ++i)
                                {
                                    Point3d pt_old = colint_left_old_perm[i];
                                    if (Functions.IsRightDirection(poly_cl, pt_old) == false)
                                    {
                                        pt_int_old_left = pt_old;
                                    }
                                }

                                Point3d pt_int_new_right = colint_right_new_perm[0];
                                Point3d pt_int_old_right = new Point3d();
                                for (int i = 0; i < colint_right_old_perm.Count; ++i)
                                {
                                    Point3d pt_old = colint_right_old_perm[i];
                                    if (Functions.IsRightDirection(poly_cl, pt_old) == true)
                                    {
                                        pt_int_old_right = pt_old;
                                    }
                                }


                                int index1 = -1;
                                for (int i = 0; i < poly_perm.NumberOfVertices; ++i)
                                {
                                    Point3d pt0 = poly_perm.GetPointAtParameter(i);
                                    Point3d pt_on_poly = poly_cl.GetClosestPointTo(pt0, Vector3d.ZAxis, false);
                                    double sta0 = poly_cl.GetDistAtPoint(pt_on_poly);
                                    if (sta0 < sta_start_switch)
                                    {
                                        col1.Add(poly_perm.GetPointAtParameter(i));
                                    }
                                    else
                                    {
                                        if (index1 == -1)
                                        {
                                            col1.Add(pt_int_old_left);
                                            col1.Add(pt_int_new_left);
                                            index1 = col1.Count;
                                        }
                                    }
                                }

                                int extra2 = 0;
                                for (int i = 0; i < poly_right.NumberOfVertices; ++i)
                                {
                                    Point3d pt0 = poly_right.GetPointAtParameter(i);
                                    Point3d pt_on_poly = poly_right.GetClosestPointTo(pt0, Vector3d.ZAxis, false);
                                    double sta0 = poly_right.GetDistAtPoint(pt_on_poly);
                                    if (sta0 > sta1R)
                                    {
                                        col1.Insert(index1, poly_right.GetPointAtParameter(i));
                                        ++extra2;
                                    }
                                }

                                for (int i = poly_left.NumberOfVertices - 1; i >= 0; --i)
                                {
                                    Point3d pt0 = poly_left.GetPointAtParameter(i);
                                    Point3d pt_on_poly = poly_left.GetClosestPointTo(pt0, Vector3d.ZAxis, false);
                                    double sta0 = poly_left.GetDistAtPoint(pt_on_poly);
                                    if (sta0 > sta1L)
                                    {
                                        col1.Insert(index1, poly_left.GetPointAtParameter(i));
                                        ++extra2;

                                    }

                                }

                                col1.Insert(index1 + extra2, pt_int_old_right);
                                col1.Insert(index1 + extra2, pt_int_new_right);

                                if (col1.Count > 1)
                                {
                                    Polyline poly_temp1 = new Polyline();
                                    for (int i = 0; i < col1.Count; ++i)
                                    {
                                        poly_temp1.AddVertexAt(i, new Point2d(col1[i].X, col1[i].Y), 0, 0, 0);
                                    }
                                    poly_temp1.Closed = true;
                                    poly_temp1.ColorIndex = 5;

                                    poly_temp1.Elevation = poly_cl.Elevation;
                                    BTrecord.AppendEntity(poly_temp1);
                                    Trans1.AddNewlyCreatedDBObject(poly_temp1, true);
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



        }

        [CommandMethod("ws2")]
        public void code_test_edit_perm_with_fillet()
        {
            if (ZZCommand_class.isSECURE() == false) return;


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            System.Data.DataTable dt1 = new System.Data.DataTable();

            dt1.Columns.Add("perm_right", typeof(double));
            dt1.Columns.Add("perm_left", typeof(double));



            dt1.Rows.Add();
            dt1.Rows[0][0] = 15;
            dt1.Rows[0][1] = 35;

            dt1.Rows.Add();
            dt1.Rows[1][0] = 35;
            dt1.Rows[1][1] = 15;


            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_cl;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_cl;
                        Prompt_cl = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_cl.SetRejectMessage("\nSelect a polyline!");
                        Prompt_cl.AllowNone = true;
                        Prompt_cl.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_cl = ThisDrawing.Editor.GetEntity(Prompt_cl);

                        if (Rezultat_cl.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_perm;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_perm;
                        Prompt_perm = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the permanent easement:");
                        Prompt_perm.SetRejectMessage("\nSelect a polyline!");
                        Prompt_perm.AllowNone = true;
                        Prompt_perm.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_perm = ThisDrawing.Editor.GetEntity(Prompt_perm);

                        if (Rezultat_perm.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the switching point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }




                        Polyline poly_perm = Trans1.GetObject(Rezultat_perm.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline poly_cl = Trans1.GetObject(Rezultat_cl.ObjectId, OpenMode.ForRead) as Polyline;

                        if (poly_perm != null && poly_cl != null && dt1 != null && dt1.Rows.Count > 0 && dt1.Rows[1][0] != DBNull.Value && dt1.Rows[1][1] != DBNull.Value)
                        {



                            DBObjectCollection col_offset_right = poly_cl.GetOffsetCurves(Convert.ToDouble(dt1.Rows[1][0]));
                            DBObjectCollection col_offset_left = poly_cl.GetOffsetCurves(-Convert.ToDouble(dt1.Rows[1][1]));
                            Polyline poly_right = col_offset_right[0] as Polyline;
                            Polyline poly_left = col_offset_left[0] as Polyline;

                            Point3dCollection col1 = new Point3dCollection();

                            Point3d pt1 = poly_cl.GetPointAtParameter(Math.Round(poly_cl.GetParameterAtPoint(poly_cl.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false)), 0));

                            double param_perm_old1 = Math.Round(poly_perm.GetParameterAtPoint(poly_perm.GetClosestPointTo(pt1, Vector3d.ZAxis, false)), 0);
                            Point3d pt_perm_old1 = poly_perm.GetPointAtParameter(param_perm_old1);


                            Polyline poly0 = new Polyline();
                            poly0.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                            poly0.AddVertexAt(1, new Point2d(pt_perm_old1.X, pt_perm_old1.Y), 0, 0, 0);
                            poly0.Elevation = poly_perm.Elevation;
                            poly0.TransformBy(Matrix3d.Scaling(300 / poly0.Length, poly0.EndPoint));
                            poly0.TransformBy(Matrix3d.Scaling(600 / poly0.Length, poly0.StartPoint));

                            Point3dCollection colint_for_left_right_side = Functions.Intersect_on_both_operands(poly0, poly_perm);
                            double param_left_perm_old = -1;
                            double param_right_perm_old = -1;

                            if (colint_for_left_right_side.Count == 2)
                            {
                                double param_perm_temp1 = Math.Round(poly_perm.GetParameterAtPoint(colint_for_left_right_side[0]), 0);
                                double param_perm_temp2 = Math.Round(poly_perm.GetParameterAtPoint(colint_for_left_right_side[1]), 0);

                                if (param_perm_temp1 < param_perm_temp2)
                                {
                                    param_left_perm_old = param_perm_temp1;
                                    param_right_perm_old = param_perm_temp2;
                                }
                                else
                                {
                                    param_left_perm_old = param_perm_temp2;
                                    param_right_perm_old = param_perm_temp1;
                                }
                            }
                            if (param_left_perm_old != -1 && param_right_perm_old != -1)
                            {


                                Point3dCollection colint_left1 = Functions.Intersect_on_both_operands(poly0, poly_left);

                                Point3d pt_perm_L2 = poly_perm.GetPointAtParameter(param_left_perm_old);
                                double param_perm_old_left1 = param_left_perm_old - 1;
                                Point3d pt_perm_L1 = poly_perm.GetPointAtParameter(param_perm_old_left1);

                                Polyline poly01L = new Polyline();
                                poly01L.AddVertexAt(0, new Point2d(pt_perm_L1.X, pt_perm_L1.Y), 0, 0, 0);
                                poly01L.AddVertexAt(1, new Point2d(pt_perm_L2.X, pt_perm_L2.Y), 0, 0, 0);
                                poly01L.Elevation = poly_perm.Elevation;

                                double param_new_L2 = Math.Round(poly_left.GetParameterAtPoint(colint_left1[0]), 0);
                                Point3d pt_newL2 = poly_left.GetPointAtParameter(param_new_L2);
                                double param_new_L3 = param_new_L2 + 1;
                                Point3d pt_newL3 = poly_left.GetPointAtParameter(param_new_L3);

                                Polyline poly02L = new Polyline();
                                poly02L.AddVertexAt(0, new Point2d(pt_newL2.X, pt_newL2.Y), 0, 0, 0);
                                poly02L.AddVertexAt(1, new Point2d(pt_newL3.X, pt_newL3.Y), 0, 0, 0);
                                poly02L.Elevation = poly_perm.Elevation;

                                Point3dCollection colint_left = Functions.Intersect_with_extend_both(poly01L, poly02L);




                                Point3dCollection colint_right1 = Functions.Intersect_on_both_operands(poly0, poly_right);

                                Point3d pt_perm_R2 = poly_perm.GetPointAtParameter(param_right_perm_old);
                                double param_perm_old_right3 = param_right_perm_old + 1;
                                Point3d pt_perm_R3 = poly_perm.GetPointAtParameter(param_perm_old_right3);

                                Polyline poly01R = new Polyline();
                                poly01R.AddVertexAt(0, new Point2d(pt_perm_R3.X, pt_perm_R3.Y), 0, 0, 0);
                                poly01R.AddVertexAt(1, new Point2d(pt_perm_R2.X, pt_perm_R2.Y), 0, 0, 0);
                                poly01R.Elevation = poly_perm.Elevation;

                                double param_new_R2 = Math.Round(poly_right.GetParameterAtPoint(colint_right1[0]), 0);
                                Point3d pt_newR2 = poly_right.GetPointAtParameter(param_new_R2);
                                double param_new_R3 = param_new_R2 + 1;
                                Point3d pt_newR3 = poly_right.GetPointAtParameter(param_new_R3);

                                Polyline poly02R = new Polyline();
                                poly02R.AddVertexAt(0, new Point2d(pt_newR2.X, pt_newR2.Y), 0, 0, 0);
                                poly02R.AddVertexAt(1, new Point2d(pt_newR3.X, pt_newR3.Y), 0, 0, 0);
                                poly02R.Elevation = poly_perm.Elevation;



                                Point3dCollection colint_right = Functions.Intersect_with_extend_both(poly01R, poly02R);


                                if (poly_left != null && poly_right != null && colint_left.Count > 0 && colint_right.Count > 0)
                                {


                                    Point3d pt_int_new_left = colint_left[0];
                                    Point3d pt_int_new_right = colint_right[0];



                                    for (int i = 0; i < poly_perm.NumberOfVertices; ++i)
                                    {
                                        if (i <= param_left_perm_old - 1)
                                        {
                                            col1.Add(poly_perm.GetPointAtParameter(i));
                                        }
                                    }

                                    col1.Add(pt_int_new_left);


                                    for (int i = 0; i < poly_left.NumberOfVertices; ++i)
                                    {
                                        if (i >= param_new_L3)
                                        {
                                            col1.Add(poly_left.GetPointAtParameter(i));


                                        }

                                    }

                                    for (int i = poly_right.NumberOfVertices - 1; i >= 0; --i)
                                    {
                                        if (i >= param_new_R3)
                                        {
                                            col1.Add(poly_right.GetPointAtParameter(i));
                                        }

                                    }

                                    col1.Add(pt_int_new_right);

                                    for (int i = 0; i < poly_perm.NumberOfVertices; ++i)
                                    {
                                        if (i >= param_right_perm_old + 1)
                                        {
                                            col1.Add(poly_perm.GetPointAtParameter(i));
                                        }
                                    }


                                    if (col1.Count > 1)
                                    {
                                        Polyline poly_temp1 = new Polyline();
                                        for (int i = 0; i < col1.Count; ++i)
                                        {
                                            poly_temp1.AddVertexAt(i, new Point2d(col1[i].X, col1[i].Y), 0, 0, 0);
                                        }
                                        poly_temp1.Closed = true;
                                        poly_temp1.ColorIndex = 5;

                                        poly_temp1.Elevation = poly_cl.Elevation;
                                        BTrecord.AppendEntity(poly_temp1);
                                        Trans1.AddNewlyCreatedDBObject(poly_temp1, true);
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



        }

        [CommandMethod("ws3")]
        public void code_test_add_tws()
        {
            if (ZZCommand_class.isSECURE() == false) return;

            System.Data.DataTable dt1 = new System.Data.DataTable();


            dt1.Columns.Add("tws_left", typeof(double));
            dt1.Columns.Add("tws_right", typeof(double));



            dt1.Rows.Add();
            dt1.Rows[0][0] = 15;
            dt1.Rows[0][1] = 5;

            dt1.Rows.Add();
            dt1.Rows[1][0] = 10;
            dt1.Rows[1][1] = 15;



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

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_cl;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_cl;
                        Prompt_cl = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_cl.SetRejectMessage("\nSelect a polyline!");
                        Prompt_cl.AllowNone = true;
                        Prompt_cl.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_cl = ThisDrawing.Editor.GetEntity(Prompt_cl);

                        if (Rezultat_cl.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_perm;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_perm;
                        Prompt_perm = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the permanent easement:");
                        Prompt_perm.SetRejectMessage("\nSelect a polyline!");
                        Prompt_perm.AllowNone = true;
                        Prompt_perm.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_perm = ThisDrawing.Editor.GetEntity(Prompt_perm);

                        if (Rezultat_perm.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Autodesk.AutoCAD.EditorInput.PromptPointResult pp_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions ppp_opt1;
                        ppp_opt1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the start point");
                        ppp_opt1.AllowNone = false;
                        pp_res1 = Editor1.GetPoint(ppp_opt1);

                        if (pp_res1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Autodesk.AutoCAD.EditorInput.PromptPointResult pp_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions ppp_opt2;
                        ppp_opt2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the end point");
                        ppp_opt2.AllowNone = false;
                        ppp_opt2.UseBasePoint = true;
                        ppp_opt2.BasePoint = pp_res1.Value;
                        pp_res2 = Editor1.GetPoint(ppp_opt2);

                        if (pp_res2.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Polyline poly_perm = Trans1.GetObject(Rezultat_perm.ObjectId, OpenMode.ForRead) as Polyline;
                        Polyline poly_cl = Trans1.GetObject(Rezultat_cl.ObjectId, OpenMode.ForRead) as Polyline;

                        if (poly_perm != null && poly_cl != null && dt1 != null && dt1.Rows.Count > 0 && dt1.Rows[0][0] != DBNull.Value && dt1.Rows[1][1] != DBNull.Value)
                        {


                            #region left
                            DBObjectCollection col_offset_left1 = poly_perm.GetOffsetCurves(-Convert.ToDouble(dt1.Rows[1][0]));
                            if (col_offset_left1.Count == 0)
                            {
                                MessageBox.Show("no left offset polyline");
                                return;
                            }

                            Polyline poly_left = col_offset_left1[0] as Polyline;

                            if (poly_left == null)
                            {
                                MessageBox.Show("no left offset polyline");
                                return;
                            }



                            #endregion

                            #region right
                            DBObjectCollection col_offset_right = poly_perm.GetOffsetCurves(-Convert.ToDouble(dt1.Rows[1][1]));
                            if (col_offset_right.Count == 0)
                            {
                                MessageBox.Show("no right offset polyline");
                                return;
                            }

                            Polyline poly_right = col_offset_right[0] as Polyline;

                            if (poly_right == null)
                            {
                                MessageBox.Show("no right offset polyline");
                                return;
                            }
                            #endregion

                            Point3dCollection col_left_tws = new Point3dCollection();
                            Point3dCollection col_right_tws = new Point3dCollection();

                            double param_cl_1 = poly_cl.GetParameterAtPoint(poly_cl.GetClosestPointTo(pp_res1.Value, Vector3d.ZAxis, false));
                            double param_cl_2 = poly_cl.GetParameterAtPoint(poly_cl.GetClosestPointTo(pp_res2.Value, Vector3d.ZAxis, false));

                            if(param_cl_1>param_cl_2)
                            {
                                double t = param_cl_1;
                                param_cl_1 = param_cl_2;
                                param_cl_2 = t;
                            }

                            if (Math.Abs(param_cl_2-param_cl_1) <0.01)
                            {
                                MessageBox.Show("points picked too close");
                                return;
                            }

                            Point3d pt1 = poly_cl.GetPointAtParameter(param_cl_1);
                            Point3d pt2 = poly_cl.GetPointAtParameter(param_cl_2);

                            #region permanent easement
                            double param_perm_1 = poly_perm.GetParameterAtPoint(poly_perm.GetClosestPointTo(pt1, Vector3d.ZAxis, false));
                            Point3d pt_perm_1 = poly_perm.GetPointAtParameter(param_perm_1);
                         
                            double param_perm_2 = poly_perm.GetParameterAtPoint(poly_perm.GetClosestPointTo(pt2, Vector3d.ZAxis, false));
                            Point3d pt_perm_2 = poly_perm.GetPointAtParameter(param_perm_2);

                            Polyline poly1 = new Polyline();
                            poly1.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                            poly1.AddVertexAt(1, new Point2d(pt_perm_1.X, pt_perm_1.Y), 0, 0, 0);
                            poly1.Elevation = poly_perm.Elevation;
                            poly1.TransformBy(Matrix3d.Scaling(300 / poly1.Length, pt_perm_1));
                            Point3dCollection colint_for_perm_easement1 = Functions.Intersect_on_both_operands(poly1, poly_perm);

                            Polyline poly2 = new Polyline();
                            poly2.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                            poly2.AddVertexAt(1, new Point2d(pt_perm_2.X, pt_perm_2.Y), 0, 0, 0);
                            poly2.Elevation = poly_perm.Elevation;
                            poly2.TransformBy(Matrix3d.Scaling(300 / poly2.Length, pt_perm_2));
                            Point3dCollection colint_for_perm_easement2 = Functions.Intersect_on_both_operands(poly2, poly_perm);


                            double param_perm_left1 = -1;
                            double param_perm_right1 = -1;

                            if (colint_for_perm_easement1.Count == 2)
                            {
                                double param_perm_temp1 = poly_perm.GetParameterAtPoint(colint_for_perm_easement1[0]);
                                double param_perm_temp2 = poly_perm.GetParameterAtPoint(colint_for_perm_easement1[1]);

                                if (param_perm_temp1 < param_perm_temp2)
                                {
                                    param_perm_left1 = param_perm_temp1;
                                    param_perm_right1 = param_perm_temp2;
                                }
                                else
                                {
                                    param_perm_left1 = param_perm_temp2;
                                    param_perm_right1 = param_perm_temp1;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly1", 9, false);
                                BTrecord.AppendEntity(poly1);
                                poly1.Layer = "_debug_poly1";
                                poly1.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly1, true);
                                Trans1.Commit();
                                return;
                            }





                            double param_perm_left2 = -1;
                            double param_perm_right2 = -1;

                            if (colint_for_perm_easement2.Count == 2)
                            {
                                double param_perm_temp1 = poly_perm.GetParameterAtPoint(colint_for_perm_easement2[0]);
                                double param_perm_temp2 = poly_perm.GetParameterAtPoint(colint_for_perm_easement2[1]);

                                if (param_perm_temp1 < param_perm_temp2)
                                {
                                    param_perm_left2 = param_perm_temp1;
                                    param_perm_right2 = param_perm_temp2;
                                }
                                else
                                {
                                    param_perm_left2 = param_perm_temp2;
                                    param_perm_right2 = param_perm_temp1;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly2", 9, false);
                                BTrecord.AppendEntity(poly2);
                                poly2.Layer = "_debug_poly2";
                                poly2.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly2, true);
                                Trans1.Commit();
                                return;
                            }

                            #endregion


                            #region left offset easement

                            double param_offset_1L = poly_left.GetParameterAtPoint(poly_left.GetClosestPointTo(pt1, Vector3d.ZAxis, false));
                            Point3d pt_offset_1L = poly_left.GetPointAtParameter(param_offset_1L);

                            Polyline poly11L = new Polyline();
                            poly11L.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                            poly11L.AddVertexAt(1, new Point2d(pt_offset_1L.X, pt_offset_1L.Y), 0, 0, 0);
                            poly11L.Elevation = poly_left.Elevation;
                            poly11L.TransformBy(Matrix3d.Scaling(300 / poly11L.Length, pt_offset_1L));

                            Point3dCollection colint_for_left_side11 = Functions.Intersect_on_both_operands(poly11L, poly_left);


                            double param_offset_left1 = -1;
                            if (colint_for_left_side11.Count == 2)
                            {
                                double param_offset_temp1 = poly_left.GetParameterAtPoint(colint_for_left_side11[0]);
                                double param_offset_temp2 = poly_left.GetParameterAtPoint(colint_for_left_side11[1]);

                                if (param_offset_temp1 < param_offset_temp2)
                                {
                                    param_offset_left1 = param_offset_temp1;
                                }
                                else
                                {
                                    param_offset_left1 = param_offset_temp2;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly_left11", 9, false);
                                BTrecord.AppendEntity(poly11L);
                                poly11L.Layer = "_debug_poly_left11";
                                poly11L.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly11L, true);
                                Trans1.Commit();
                                return;
                            }


                            double param_offset_2L = poly_left.GetParameterAtPoint(poly_left.GetClosestPointTo(pt2, Vector3d.ZAxis, false));
                            Point3d pt_offset_2L = poly_left.GetPointAtParameter(param_offset_2L);

                            Polyline poly22L = new Polyline();
                            poly22L.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                            poly22L.AddVertexAt(1, new Point2d(pt_offset_2L.X, pt_offset_2L.Y), 0, 0, 0);
                            poly22L.Elevation = poly_left.Elevation;
                            poly22L.TransformBy(Matrix3d.Scaling(300 / poly22L.Length, pt_offset_2L));

                            Point3dCollection colint_for_left_side22 = Functions.Intersect_on_both_operands(poly22L, poly_left);
                            double param_offset_left2 = -1;

                            if (colint_for_left_side22.Count == 2)
                            {
                                double param_offset_temp1 = poly_left.GetParameterAtPoint(colint_for_left_side22[0]);
                                double param_offset_temp2 = poly_left.GetParameterAtPoint(colint_for_left_side22[1]);

                                if (param_offset_temp1 < param_offset_temp2)
                                {
                                    param_offset_left2 = param_offset_temp1;
                                }
                                else
                                {
                                    param_offset_left2 = param_offset_temp2;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly_left22", 9, false);
                                BTrecord.AppendEntity(poly22L);
                                poly22L.Layer = "_debug_poly_left22";
                                poly22L.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly22L, true);
                                Trans1.Commit();
                                return;
                            }

                            #endregion

                            #region right offset easement

                            double param_offset_1R = poly_right.GetParameterAtPoint(poly_right.GetClosestPointTo(pt1, Vector3d.ZAxis, false));
                            Point3d pt_offset_1R = poly_right.GetPointAtParameter(param_offset_1R);

                            Polyline poly11R = new Polyline();
                            poly11R.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                            poly11R.AddVertexAt(1, new Point2d(pt_offset_1R.X, pt_offset_1R.Y), 0, 0, 0);
                            poly11R.Elevation = poly_right.Elevation;
                            poly11R.TransformBy(Matrix3d.Scaling(300 / poly11R.Length, pt_offset_1R));

                            Point3dCollection colint_for_right_side11 = Functions.Intersect_on_both_operands(poly11R, poly_right);


                            double param_offset_right1 = -1;
                            if (colint_for_right_side11.Count == 2)
                            {
                                double param_offset_temp1 = poly_right.GetParameterAtPoint(colint_for_right_side11[0]);
                                double param_offset_temp2 = poly_right.GetParameterAtPoint(colint_for_right_side11[1]);

                                if (param_offset_temp1 > param_offset_temp2)
                                {
                                    param_offset_right1 = param_offset_temp1;
                                }
                                else
                                {
                                    param_offset_right1 = param_offset_temp2;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly_right11", 9, false);
                                BTrecord.AppendEntity(poly11R);
                                poly11R.Layer = "_debug_poly_right11";
                                poly11R.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly11R, true);
                                Trans1.Commit();
                                return;
                            }



                            double param_offset_2R = poly_right.GetParameterAtPoint(poly_right.GetClosestPointTo(pt2, Vector3d.ZAxis, false));
                            Point3d pt_offset_2R = poly_right.GetPointAtParameter(param_offset_2R);

                            Polyline poly22R = new Polyline();
                            poly22R.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                            poly22R.AddVertexAt(1, new Point2d(pt_offset_2R.X, pt_offset_2R.Y), 0, 0, 0);
                            poly22R.Elevation = poly_right.Elevation;
                            poly22R.TransformBy(Matrix3d.Scaling(300 / poly22R.Length, pt_offset_2R));

                            Point3dCollection colint_for_right_side22 = Functions.Intersect_on_both_operands(poly22R, poly_right);
                            double param_offset_right2 = -1;

                            if (colint_for_right_side22.Count == 2)
                            {
                                double param_offset_temp1 = poly_right.GetParameterAtPoint(colint_for_right_side22[0]);
                                double param_offset_temp2 = poly_right.GetParameterAtPoint(colint_for_right_side22[1]);

                                if (param_offset_temp1 > param_offset_temp2)
                                {
                                    param_offset_right2 = param_offset_temp1;
                                }
                                else
                                {
                                    param_offset_right2 = param_offset_temp2;
                                }
                            }
                            else
                            {
                                Functions.Creaza_layer("_debug_poly_right22", 9, false);
                                BTrecord.AppendEntity(poly22R);
                                poly22R.Layer = "_debug_poly_right22";
                                poly22R.ColorIndex = 256;
                                Trans1.AddNewlyCreatedDBObject(poly22R, true);
                                Trans1.Commit();
                                return;
                            }

                            #endregion

                            if (param_perm_left1 != -1 && param_perm_left2 != -1 && param_offset_left1 != -1 && param_offset_left2 != -1)
                            {

                                List<Point3d> lista_puncte = new List<Point3d>();

                                col_left_tws.Add(poly_left.GetPointAtParameter(param_offset_left1));

                                for (int i = 0; i < poly_left.NumberOfVertices; ++i)
                                {


                                    if (i > param_offset_left1 && i < param_offset_left2)
                                    {
                                        col_left_tws.Add(poly_left.GetPointAtParameter(i));
                                    }
                                }

                                col_left_tws.Add(poly_left.GetPointAtParameter(param_offset_left2));

                                col_left_tws.Add(poly_perm.GetPointAtParameter(param_perm_left2));



                                for (int i = poly_perm.NumberOfVertices - 1; i >= 0; --i)
                                {
                                    if (i > param_perm_left1 && i < param_perm_left2)
                                    {
                                        col_left_tws.Add(poly_perm.GetPointAtParameter(i));
                                    }
                                }
                                col_left_tws.Add(poly_perm.GetPointAtParameter(param_perm_left1));

                                if (col_left_tws.Count > 1)
                                {
                                    Polyline poly_temp1 = new Polyline();
                                    for (int i = 0; i < col_left_tws.Count; ++i)
                                    {
                                        poly_temp1.AddVertexAt(i, new Point2d(col_left_tws[i].X, col_left_tws[i].Y), 0, 0, 0);
                                    }
                                    poly_temp1.Closed = true;
                                    poly_temp1.ColorIndex = 3;

                                    poly_temp1.Elevation = poly_cl.Elevation;
                                    BTrecord.AppendEntity(poly_temp1);
                                    Trans1.AddNewlyCreatedDBObject(poly_temp1, true);
                                }

                            }

                            if (param_perm_right1 != -1 && param_perm_right2 != -1 && param_offset_right1 != -1 && param_offset_right2 != -1)
                            {
                                List<Point3d> lista_puncte = new List<Point3d>();



                                col_right_tws.Add(poly_perm.GetPointAtParameter(param_perm_right1));
                                for (int i = poly_perm.NumberOfVertices - 1; i >= 0; --i)
                                {
                                    if (i > param_perm_right2 && i < param_perm_right1)
                                    {
                                        col_right_tws.Add(poly_perm.GetPointAtParameter(i));
                                    }
                                }
                                col_right_tws.Add(poly_perm.GetPointAtParameter(param_perm_right2));

                                col_right_tws.Add(poly_right.GetPointAtParameter(param_offset_right2));
                                for (int i = 0; i < poly_right.NumberOfVertices; ++i)
                                {
                                    if (i > param_offset_right2 && i < param_offset_right1)
                                    {
                                        col_right_tws.Add(poly_right.GetPointAtParameter(i));
                                    }
                                }
                                col_right_tws.Add(poly_right.GetPointAtParameter(param_offset_right1));

                                if (col_right_tws.Count > 1)
                                {
                                    Polyline poly_temp1 = new Polyline();
                                    for (int i = 0; i < col_right_tws.Count; ++i)
                                    {
                                        poly_temp1.AddVertexAt(i, new Point2d(col_right_tws[i].X, col_right_tws[i].Y), 0, 0, 0);
                                    }
                                    poly_temp1.Closed = true;
                                    poly_temp1.ColorIndex = 2;

                                    poly_temp1.Elevation = poly_cl.Elevation;
                                    BTrecord.AppendEntity(poly_temp1);
                                    Trans1.AddNewlyCreatedDBObject(poly_temp1, true);
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



        }

    }
}
