using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;


namespace Dimensioning
{
    public class Command_class
    {

        public bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "8CDA6CE3":
                    return true;
                case "36D79DE5":
                    return true;
                case "FEA3192C":
                    return true;
                case "B454BD5B":
                    return true;
                case "6E40460D":
                    return true;
                case "0892E01D":
                    return true;
                case "4ED21ABF":
                    return true;
                case "56766C69":
                    return true;
                case "DA214366":
                    return true;
                case "3CF68AF2":
                    return true;
                case "389A2249":
                    return true;
                case "AED6B68E":
                    return true;
                case "8C040338":
                    return true;
                case "8CD08F48":
                    return true;
                case "0E26E402":
                    return true;
                case "4A123A50":
                    return true;

                case "98D9B617":
                    return true;
                case "B838FEB4":
                    return true;
                case "1AE1721C":
                    return true;
                case "CA9E6FFE":
                    return true;
                case "DE281128":
                    return true;
                case "FC7C4F1":
                    return true;
                case "B67EC134":
                    return true;
                case "E64DBF0A":
                    return true;
                case "561F1509":
                    return true;
                case "B63AD3F6":
                    return true;

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case  "18399D24":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        return false;
                    }
            }
        }

        public string GetHDDSerialNumber(string drive)
        {
            //check to see if the user provided a drive letter
            //if not default it to "C"
            if (drive == "" || drive == null)
            {
                drive = "C";
            }
            //create our ManagementObject, passing it the drive letter to the
            //DevideID using WQL
            ManagementObject disk = new ManagementObject("Win32_LogicalDisk.DeviceID=\"" + drive + ":\"");
            //bind our management object
            disk.Get();
            //return the serial number
            return disk["VolumeSerialNumber"].ToString();
        }

        static public Point3dCollection IntersectOnBothOperands(Curve Curba1, Curve Curba2)
        {
            Point3dCollection Col_int = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands = new Point3dCollection();


            Curba1.IntersectWith(Curba2, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero);

            if (Col_int.Count > 0)
            {
                for (int i = 0; i < Col_int.Count; ++i)
                {
                    Point3d point_int = new Point3d();
                    point_int = Col_int[i];
                    try
                    {
                        double param_on_1;
                        param_on_1 = Curba1.GetParameterAtPoint(point_int);
                        double param_on_2;
                        param_on_2 = Curba2.GetParameterAtPoint(point_int);

                        if (Col_int_on_both_operands.Contains(point_int) == false)
                        {
                            Col_int_on_both_operands.Add(point_int);
                        }
                    }

                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                    }

                }
            }

            return Col_int_on_both_operands;
        }

        public void Creaza_layer(string Layername, short Culoare, bool Plot)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
                        LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        if (LayerTable1.Has(Layername) == false)
                        {
                            LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            new_layer.Name = Layername;
                            new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                            new_layer.IsPlottable = Plot;
                            LayerTable1.Add(new_layer);
                            Trans1.AddNewlyCreatedDBObject(new_layer, true);

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

        [CommandMethod("DIM_EX1")]
        public void CREATE_EXPLODED_DIMS_BETWEEN_TWO_POLYLINES()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters
            double Len1 = 8;
            double Width = 2.67;
            // text parameters
            double Height1 = 8;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly1;
                        Prompt_Poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the first polyline:");
                        Prompt_Poly1.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly1.AllowNone = true;
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_Poly1);

                        if (Rezultat_poly1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        //PolyCL_MS = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly2;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly2;
                        Prompt_Poly2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the second polyline:");
                        Prompt_Poly2.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly2.AllowNone = true;
                        Prompt_Poly2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly2 = ThisDrawing.Editor.GetEntity(Prompt_Poly2);

                        if (Rezultat_poly2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly3;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly3;
                        Prompt_Poly3 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the a rectangle:");
                        Prompt_Poly3.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly3.AllowNone = true;
                        Prompt_Poly3.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly3 = ThisDrawing.Editor.GetEntity(Prompt_Poly3);

                        if (Rezultat_poly2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);
                        Polyline Poly2 = (Polyline)Trans1.GetObject(Rezultat_poly2.ObjectId, OpenMode.ForRead);
                        Polyline Rectangle_viewport = (Polyline)Trans1.GetObject(Rezultat_poly3.ObjectId, OpenMode.ForRead);
                        foreach (ObjectId ObjID in BTrecord)
                        {
                            Entity Ent1 = (Entity)Trans1.GetObject(ObjID, OpenMode.ForRead);
                            if (Ent1.Layer == Rectangle_viewport.Layer)
                            {
                                if (Ent1 is Polyline)
                                {
                                    Polyline Poly3 = (Polyline)Ent1;
                                    Point3dCollection Col_int = new Point3dCollection();
                                    Col_int = IntersectOnBothOperands(Poly3, Poly1);
                                    // Poly3.IntersectWith(Poly1, Autodesk.AutoCAD.DatabaseServices.Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero);
                                    if (Col_int.Count == 2)
                                    {
                                        Double Param1 = Poly1.GetParameterAtPoint(Col_int[0]);
                                        Double Param2 = Poly1.GetParameterAtPoint(Col_int[1]);
                                        if (Param1 > Param2)
                                        {
                                            double t = Param1;
                                            Param1 = Param2;
                                            Param2 = t;
                                        }

                                        Polyline Poly_dim = new Polyline();
                                        Poly_dim.AddVertexAt(0, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                                        int iDX = 1;
                                        if (Math.Floor(Param1) < Math.Floor(Param2))
                                        {
                                            int Start1 = Convert.ToInt32(Math.Ceiling(Param1));

                                            for (int j = Start1; j <= Math.Floor(Param2); ++j)
                                            {
                                                Poly_dim.AddVertexAt(iDX, Poly1.GetPoint2dAt(j), 0, 0, 0);
                                                iDX = iDX + 1;
                                            }
                                        }
                                        Poly_dim.AddVertexAt(iDX, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                                        for (int i = 1; i < Poly_dim.NumberOfVertices; ++i)
                                        {
                                            Point3d Pt1 = Poly_dim.GetPoint3dAt(i - 1);
                                            Point3d Pt2 = Poly_dim.GetPoint3dAt(i);
                                            Point3d PtM = new Point3d((Pt1.X + Pt2.X) / 2, (Pt1.Y + Pt2.Y) / 2, 0);

                                            Line Line0 = new Line(PtM, Pt2);
                                            Line0.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PtM));
                                            Point3dCollection Col_int_temp = new Point3dCollection();
                                            Line0.IntersectWith(Poly2, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_temp, IntPtr.Zero, IntPtr.Zero);

                                            Point3d Pt_on_poly = Poly2.GetClosestPointTo(PtM, Vector3d.ZAxis, false);


                                            Line Line1 = new Line();

                                            if (Col_int_temp.Count > 0)
                                            {
                                                for (int k = 0; k < Col_int_temp.Count; ++k)
                                                {
                                                    Line Linet = new Line(PtM, Col_int_temp[k]);
                                                    if (k == 0)
                                                    {
                                                        Line1 = new Line(PtM, Col_int_temp[k]);
                                                    }
                                                    else
                                                    {
                                                        if (Line1.Length > Linet.Length)
                                                        {
                                                            Line1 = new Line(PtM, Col_int_temp[k]);
                                                        }
                                                    }
                                                }

                                                if (Line1.Length < 300)
                                                {

                                                    Point3d Punct8 = Line1.GetPointAtDist(Len1);
                                                    Point3d Punct267 = Line1.GetPointAtDist(Width);
                                                    Point3d Punct0 = Line1.StartPoint;
                                                    Line Linew = new Line(Punct267, Punct0);
                                                    Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));
                                                    Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                                                    Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));

                                                    Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                                                    BTrecord.AppendEntity(Solid1);
                                                    Trans1.AddNewlyCreatedDBObject(Solid1, true);

                                                    Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                                                    Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                                                    Point3d Punct00 = Line1.EndPoint;
                                                    Line Lineww = new Line(Punct267267, Punct00);
                                                    Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));
                                                    Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                                                    Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));

                                                    Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                                                    BTrecord.AppendEntity(Solid11);
                                                    Trans1.AddNewlyCreatedDBObject(Solid11, true);

                                                    DBText Text1 = new DBText();
                                                    Text1.Rotation = 0;
                                                    Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                                                    Text1.Justify = AttachmentPoint.MiddleCenter;
                                                    Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                                                    Text1.Height = Height1;

                                                    BTrecord.AppendEntity(Text1);
                                                    Trans1.AddNewlyCreatedDBObject(Text1, true);

                                                    BTrecord.AppendEntity(Line1);
                                                    Trans1.AddNewlyCreatedDBObject(Line1, true);
                                                }
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
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("DIM_EX2")]
        public void CREATE_manual_EXPLODED_DIMS_BETWEEN_TWO_POLYLINES()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters
            double Len1 = 8;
            double Width = 2.67;
            // text parameters
            double Height1 = 8;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly1;
                        Prompt_Poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the first polyline:");
                        Prompt_Poly1.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly1.AllowNone = true;
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_Poly1);

                        if (Rezultat_poly1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        //PolyCL_MS = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly2;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly2;
                        Prompt_Poly2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the second polyline:");
                        Prompt_Poly2.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly2.AllowNone = true;
                        Prompt_Poly2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly2 = ThisDrawing.Editor.GetEntity(Prompt_Poly2);

                        if (Rezultat_poly2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);
                        Polyline Poly2 = (Polyline)Trans1.GetObject(Rezultat_poly2.ObjectId, OpenMode.ForRead);
                        Point3d PtM = Poly1.GetClosestPointTo(Rezultat_poly1.PickedPoint, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(PtM);


                        Point3d Pt1 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Floor(Param1)));
                        Point3d Pt2 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Ceiling(Param1)));
                        Line Line0 = new Line(PtM, Pt2);
                        Line0.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PtM));

                        Point3dCollection Col_int_temp = new Point3dCollection();
                        Line0.IntersectWith(Poly2, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_temp, IntPtr.Zero, IntPtr.Zero);

                        if (Col_int_temp.Count > 0)
                        {
                            Point3d Pt_on_poly = Poly2.GetClosestPointTo(PtM, Vector3d.ZAxis, false);


                            Line Line1 = new Line();

                            for (int i = 0; i < Col_int_temp.Count; ++i)
                            {
                                Line Linet = new Line(PtM, Col_int_temp[i]);
                                if (i == 0)
                                {
                                    Line1 = new Line(PtM, Col_int_temp[i]);
                                }
                                else
                                {
                                    if (Line1.Length > Linet.Length)
                                    {
                                        Line1 = new Line(PtM, Col_int_temp[i]);
                                    }
                                }
                            }

                            Point3d Punct8 = Line1.GetPointAtDist(Len1);
                            Point3d Punct267 = Line1.GetPointAtDist(Width);
                            Point3d Punct0 = Line1.StartPoint;
                            Line Linew = new Line(Punct267, Punct0);
                            Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));
                            Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                            Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));

                            Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                            BTrecord.AppendEntity(Solid1);
                            Trans1.AddNewlyCreatedDBObject(Solid1, true);

                            Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                            Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                            Point3d Punct00 = Line1.EndPoint;
                            Line Lineww = new Line(Punct267267, Punct00);
                            Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));
                            Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                            Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));

                            Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                            BTrecord.AppendEntity(Solid11);
                            Trans1.AddNewlyCreatedDBObject(Solid11, true);

                            DBText Text1 = new DBText();
                            Text1.Rotation = 0;
                            Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                            Text1.Justify = AttachmentPoint.MiddleCenter;
                            Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                            Text1.Height = Height1;

                            BTrecord.AppendEntity(Text1);
                            Trans1.AddNewlyCreatedDBObject(Text1, true);

                            BTrecord.AppendEntity(Line1);
                            Trans1.AddNewlyCreatedDBObject(Line1, true);

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

        [CommandMethod("DIM_EX3")]
        public void CREATE_manual_multiples_EXPLODED_DIMS_BETWEEN_TWO_POLYLINES()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters
            double Len1 = 8;
            double Width = 2.67;
            // text parameters
            double Height1 = 8;

            double min_dist = 20;
            double extra_length = 8;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Polyline Poly1;
                Polyline Poly2;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly1;
                        Prompt_Poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect a point on the first polyline:");
                        Prompt_Poly1.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly1.AllowNone = true;
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_Poly1);

                        if (Rezultat_poly1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        //PolyCL_MS = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly2;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly2;
                        Prompt_Poly2 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the second polyline:");
                        Prompt_Poly2.SetRejectMessage("\nSelect a polyline!");
                        Prompt_Poly2.AllowNone = true;
                        Prompt_Poly2.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_poly2 = ThisDrawing.Editor.GetEntity(Prompt_Poly2);

                        if (Rezultat_poly2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Poly1 = (Polyline)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);
                        Poly2 = (Polyline)Trans1.GetObject(Rezultat_poly2.ObjectId, OpenMode.ForRead);
                        Point3d PtM = Poly1.GetClosestPointTo(Rezultat_poly1.PickedPoint, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(PtM);


                        Point3d Pt1 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Floor(Param1)));
                        Point3d Pt2 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Ceiling(Param1)));
                        Line Line0 = new Line(PtM, Pt2);
                        Line0.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PtM));

                        Point3dCollection Col_int_temp = new Point3dCollection();
                        Line0.IntersectWith(Poly2, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_temp, IntPtr.Zero, IntPtr.Zero);

                        if (Col_int_temp.Count > 0)
                        {
                            Point3d Pt_on_poly = Poly2.GetClosestPointTo(PtM, Vector3d.ZAxis, false);


                            Line Line1 = new Line();

                            for (int i = 0; i < Col_int_temp.Count; ++i)
                            {
                                Line Linet = new Line(PtM, Col_int_temp[i]);
                                if (i == 0)
                                {
                                    Line1 = new Line(PtM, Col_int_temp[i]);
                                }
                                else
                                {
                                    if (Line1.Length > Linet.Length)
                                    {
                                        Line1 = new Line(PtM, Col_int_temp[i]);
                                    }
                                }
                            }

                            if (Line1.Length >= min_dist)
                            {

                                Point3d Punct8 = Line1.GetPointAtDist(Len1);
                                Point3d Punct267 = Line1.GetPointAtDist(Width);
                                Point3d Punct0 = Line1.StartPoint;
                                Line Linew = new Line(Punct267, Punct0);
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));
                                Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                                Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));

                                Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                                BTrecord.AppendEntity(Solid1);
                                Trans1.AddNewlyCreatedDBObject(Solid1, true);

                                Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                                Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                                Point3d Punct00 = Line1.EndPoint;
                                Line Lineww = new Line(Punct267267, Punct00);
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));
                                Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                                Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));

                                Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                                BTrecord.AppendEntity(Solid11);
                                Trans1.AddNewlyCreatedDBObject(Solid11, true);

                                DBText Text1 = new DBText();
                                Text1.Rotation = 0;
                                Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                                Text1.Justify = AttachmentPoint.MiddleCenter;
                                Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                                Text1.Height = Height1;

                                BTrecord.AppendEntity(Text1);
                                Trans1.AddNewlyCreatedDBObject(Text1, true);

                                BTrecord.AppendEntity(Line1);
                                Trans1.AddNewlyCreatedDBObject(Line1, true);
                            }
                            else
                            {
                                Point3d Punct8 = Line1.GetPointAtDist(Len1);
                                Point3d Punct267 = Line1.GetPointAtDist(Width);
                                Point3d Punct0 = Line1.StartPoint;
                                Line Linew = new Line(Punct267, Punct0);
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));



                                Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                                Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Punct0));

                                Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                                BTrecord.AppendEntity(Solid1);
                                Trans1.AddNewlyCreatedDBObject(Solid1, true);

                                Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                                Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                                Point3d Punct00 = Line1.EndPoint;
                                Line Lineww = new Line(Punct267267, Punct00);
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));



                                Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                                Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Punct00));

                                Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                                BTrecord.AppendEntity(Solid11);
                                Trans1.AddNewlyCreatedDBObject(Solid11, true);

                                DBText Text1 = new DBText();
                                Text1.Rotation = 0;
                                Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                                Text1.Justify = AttachmentPoint.MiddleCenter;
                                Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                                Text1.Height = Height1;

                                BTrecord.AppendEntity(Text1);
                                Trans1.AddNewlyCreatedDBObject(Text1, true);


                                double Scale1 = (Line1.Length + 2 * extra_length + 2 * Len1) / Line1.Length;
                                Line1.TransformBy(Matrix3d.Scaling(Scale1, Line1.GetPointAtDist(Line1.Length / 2)));
                                BTrecord.AppendEntity(Line1);
                                Trans1.AddNewlyCreatedDBObject(Line1, true);
                            }






                        }




                        Trans1.Commit();
                    }

                start1:

                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the next point");
                    PP1.AllowNone = false;
                    Point_res1 = Editor1.GetPoint(PP1);

                    if (Point_res1.Status != PromptStatus.OK)
                    {

                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        return;
                    }

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Point3d PtM = Poly1.GetClosestPointTo(Point_res1.Value, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(PtM);


                        Point3d Pt1 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Floor(Param1)));
                        Point3d Pt2 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Ceiling(Param1)));
                        Line Line0 = new Line(PtM, Pt2);
                        Line0.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PtM));

                        Point3dCollection Col_int_temp = new Point3dCollection();
                        Line0.IntersectWith(Poly2, Autodesk.AutoCAD.DatabaseServices.Intersect.ExtendThis, Col_int_temp, IntPtr.Zero, IntPtr.Zero);

                        if (Col_int_temp.Count > 0)
                        {
                            Point3d Pt_on_poly = Poly2.GetClosestPointTo(PtM, Vector3d.ZAxis, false);

                            Line Line1 = new Line();

                            for (int i = 0; i < Col_int_temp.Count; ++i)
                            {
                                Line Linet = new Line(PtM, Col_int_temp[i]);
                                if (i == 0)
                                {
                                    Line1 = new Line(PtM, Col_int_temp[i]);
                                }
                                else
                                {
                                    if (Line1.Length > Linet.Length)
                                    {
                                        Line1 = new Line(PtM, Col_int_temp[i]);
                                    }
                                }
                            }


                            if (Line1.Length >= min_dist)
                            {

                                Point3d Punct8 = Line1.GetPointAtDist(Len1);
                                Point3d Punct267 = Line1.GetPointAtDist(Width);
                                Point3d Punct0 = Line1.StartPoint;
                                Line Linew = new Line(Punct267, Punct0);
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));
                                Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                                Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));

                                Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                                BTrecord.AppendEntity(Solid1);
                                Trans1.AddNewlyCreatedDBObject(Solid1, true);

                                Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                                Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                                Point3d Punct00 = Line1.EndPoint;
                                Line Lineww = new Line(Punct267267, Punct00);
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));
                                Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                                Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));

                                Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                                BTrecord.AppendEntity(Solid11);
                                Trans1.AddNewlyCreatedDBObject(Solid11, true);

                                DBText Text1 = new DBText();
                                Text1.Rotation = 0;
                                Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                                Text1.Justify = AttachmentPoint.MiddleCenter;
                                Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                                Text1.Height = Height1;

                                BTrecord.AppendEntity(Text1);
                                Trans1.AddNewlyCreatedDBObject(Text1, true);

                                BTrecord.AppendEntity(Line1);
                                Trans1.AddNewlyCreatedDBObject(Line1, true);
                            }
                            else
                            {
                                Point3d Punct8 = Line1.GetPointAtDist(Len1);
                                Point3d Punct267 = Line1.GetPointAtDist(Width);
                                Point3d Punct0 = Line1.StartPoint;
                                Line Linew = new Line(Punct267, Punct0);
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267));



                                Point3d Punctm = new Point3d((Linew.StartPoint.X + Linew.EndPoint.X) / 2, (Linew.StartPoint.Y + Linew.EndPoint.Y) / 2, 0);
                                Linew.TransformBy(Matrix3d.Displacement(Punctm.GetVectorTo(Punct8)));
                                Linew.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Punct0));

                                Solid Solid1 = new Solid(Linew.StartPoint, Linew.EndPoint, Punct0);
                                BTrecord.AppendEntity(Solid1);
                                Trans1.AddNewlyCreatedDBObject(Solid1, true);

                                Point3d Punct88 = Line1.GetPointAtDist(Line1.Length - Len1);
                                Point3d Punct267267 = Line1.GetPointAtDist(Line1.Length - Width);
                                Point3d Punct00 = Line1.EndPoint;
                                Line Lineww = new Line(Punct267267, Punct00);
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Punct267267));



                                Point3d Punctmm = new Point3d((Lineww.StartPoint.X + Lineww.EndPoint.X) / 2, (Lineww.StartPoint.Y + Lineww.EndPoint.Y) / 2, 0);
                                Lineww.TransformBy(Matrix3d.Displacement(Punctmm.GetVectorTo(Punct88)));
                                Lineww.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Punct00));

                                Solid Solid11 = new Solid(Lineww.StartPoint, Lineww.EndPoint, Punct00);

                                BTrecord.AppendEntity(Solid11);
                                Trans1.AddNewlyCreatedDBObject(Solid11, true);

                                DBText Text1 = new DBText();
                                Text1.Rotation = 0;
                                Text1.TextString = Convert.ToString(Math.Round(Line1.Length, 0)) + "'";
                                Text1.Justify = AttachmentPoint.MiddleCenter;
                                Text1.AlignmentPoint = new Point3d((Line1.StartPoint.X + Line1.EndPoint.X) / 2, (Line1.StartPoint.Y + Line1.EndPoint.Y) / 2, 0);

                                Text1.Height = Height1;

                                BTrecord.AppendEntity(Text1);
                                Trans1.AddNewlyCreatedDBObject(Text1, true);


                                double Scale1 = (Line1.Length + 2 * extra_length + 2 * Len1) / Line1.Length;
                                Line1.TransformBy(Matrix3d.Scaling(Scale1, Line1.GetPointAtDist(Line1.Length / 2)));
                                BTrecord.AppendEntity(Line1);
                                Trans1.AddNewlyCreatedDBObject(Line1, true);
                            }




                        }




                        Trans1.Commit();
                    }
                    goto start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("remove_and_add_suffix")]
        public void remove_and_add_suffix()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters
            string feet1 = "'";


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the text or Mtext objects:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is DBText)
                            {
                                DBText Text1 = (DBText)Ent1;
                                Text1.UpgradeOpen();
                                Text1.TextString = Text1.TextString.Replace("'", "");
                                Text1.TextString = Text1.TextString + "'";

                            }
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;
                                Mtext1.UpgradeOpen();
                                Mtext1.Contents = Mtext1.Contents.Replace("'", "");
                                Mtext1.Contents = Mtext1.Contents + "'";

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

        [CommandMethod("REPLACE_PL_EXCLAM", CommandFlags.UsePickSet)]
        public void REPLACE_CL_WITH_EX()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the text or Mtext objects:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;

                                string NEW_TXT1 = Mtext1.Contents;

                                if (Mtext1.Contents.Contains("⅊") == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    string NEW_TXT = "{\fromans|c0;!}";
                                    int ii = 92;
                                    char cc = (char)ii;
                                    NEW_TXT = "{" + cc.ToString() + "Fromans|c0;!}";
                                    Mtext1.Contents = NEW_TXT;

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
        }

        [CommandMethod("REPLACE_X_ING", CommandFlags.UsePickSet)]
        public void REPLACE_xing_WITH_nothing()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the Mtext objects:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;

                                string NEW_TXT1 = Mtext1.Contents;

                                if (Mtext1.Contents.Contains(" X-ING") == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    Mtext1.Contents = Mtext1.Contents.Replace(" X-ING", "");

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
        }

        [CommandMethod("DIM_DIM1")]
        public void CREATE_DIMS_BETWEEN_TWO_POints()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters

            // text parameters

            ObjectId DimStyleID1 = ObjectId.Null;
            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                Creaza_layer("TEXT", 2, true);

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.TextStyleTable Text_style_table = (TextStyleTable)Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.DimStyleTable Dim_style_table = (DimStyleTable)Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        Boolean Exista = false;
                        TextStyleTableRecord Text_style_romans;


                        foreach (ObjectId TextStyle_id in Text_style_table)
                        {
                            TextStyleTableRecord TextStyle = (TextStyleTableRecord)Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (TextStyle.Name == "Romans_DimStyle")
                            {
                                TextStyle.UpgradeOpen();
                                TextStyle.TextSize = 0;
                                TextStyle.FileName = "romans.shx";
                                TextStyle.TextSize = 0;
                                TextStyle.ObliquingAngle = 0;
                                TextStyle.XScale = 1.0;

                                Text_style_romans = TextStyle;
                                Exista = true;
                                TextStyleID1 = TextStyle.ObjectId;
                                goto Label1;
                            }
                        }


                    Label1:


                        if (Exista == false)
                        {
                            Text_style_table.UpgradeOpen();
                            Text_style_romans = new TextStyleTableRecord();
                            Text_style_romans.Name = "Romans_DimStyle";

                            Text_style_romans.TextSize = 0;
                            Text_style_romans.ObliquingAngle = 0;
                            Text_style_romans.FileName = "romans.shx";
                            Text_style_romans.XScale = 1.0;
                            Text_style_table.Add(Text_style_romans);
                            Trans1.AddNewlyCreatedDBObject(Text_style_romans, true);
                            TextStyleID1 = Text_style_romans.ObjectId;
                        }


                        Boolean Exista1 = false;
                        DimStyleTableRecord Dim_style_Wksp_dim;


                        foreach (ObjectId DimStyle_id in Dim_style_table)
                        {
                            DimStyleTableRecord DimStyle1 = (DimStyleTableRecord)Trans1.GetObject(DimStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (DimStyle1.Name == "Wksp Dim")
                            {
                                DimStyle1.UpgradeOpen();
                                DimStyle1.Dimasz = 24;
                                DimStyle1.Dimtxt = 16;
                                DimStyle1.Dimtxtdirection = false;
                                DimStyle1.Dimtxsty = TextStyleID1;
                                DimStyle1.Dimdec = 0;
                                DimStyle1.Dimtih = false;
                                DimStyle1.Dimtoh = false;
                                DimStyle1.Dimpost = "'";
                                DimStyle1.Dimse1 = true;
                                DimStyle1.Dimse2 = true;
                                DimStyle1.Dimcen = 0;
                                DimStyle1.Dimgap = 6;
                                DimStyle1.Dimatfit = 0;

                                Dim_style_Wksp_dim = DimStyle1;
                                Exista1 = true;
                                DimStyleID1 = DimStyle1.ObjectId;
                                goto Label2;
                            }
                        }


                    Label2:

                        if (Exista1 == false)
                        {
                            Dim_style_table.UpgradeOpen();
                            Dim_style_Wksp_dim = new DimStyleTableRecord();

                            Dim_style_Wksp_dim.Name = "Wksp Dim";
                            Dim_style_Wksp_dim.Dimasz = 24;
                            Dim_style_Wksp_dim.Dimtxt = 16;
                            Dim_style_Wksp_dim.Dimtxtdirection = false;
                            Dim_style_Wksp_dim.Dimtxsty = TextStyleID1;
                            Dim_style_Wksp_dim.Dimdec = 0;
                            Dim_style_Wksp_dim.Dimtih = false;
                            Dim_style_Wksp_dim.Dimtoh = false;
                            Dim_style_Wksp_dim.Dimpost = "'";
                            Dim_style_Wksp_dim.Dimse1 = true;
                            Dim_style_Wksp_dim.Dimse2 = true;
                            Dim_style_Wksp_dim.Dimcen = 0;
                            Dim_style_Wksp_dim.Dimgap = 6;
                            Dim_style_Wksp_dim.Dimatfit = 0;

                            Dim_style_table.Add(Dim_style_Wksp_dim);
                            Trans1.AddNewlyCreatedDBObject(Dim_style_Wksp_dim, true);
                            DimStyleID1 = Dim_style_Wksp_dim.ObjectId;
                        }

                        Trans1.Commit();

                    }

                Start1:

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);




                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                


                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");

                        object NEW_OSnap = 512;


                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        NEW_OSnap = 128;
                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);


                        AlignedDimension Dimension1 = new AlignedDimension();
                        Dimension1.XLine1Point = Punct1;
                        Dimension1.XLine2Point = Punct2;

                        Dimension1.DimLinePoint = Punct1;
                        Dimension1.DimensionStyle = DimStyleID1;
                        Dimension1.Layer = "TEXT";

                        BTrecord.AppendEntity(Dimension1);
                        Trans1.AddNewlyCreatedDBObject(Dimension1, true);







                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                        Trans1.Commit();
                    }
                    goto Start1;

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("DIM_atws")]
        public void CREATE_mtext_rounded_for_ATWS()
        {
            if (isSECURE() == false)
            {
                return;
            }


            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                Creaza_layer("TEXT", 2, true);

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.DatabaseServices.TextStyleTable Text_style_table = (TextStyleTable)Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.DimStyleTable Dim_style_table = (DimStyleTable)Trans1.GetObject(ThisDrawing.Database.DimStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                        Boolean Exista = false;
                        TextStyleTableRecord Text_style_romans;


                        foreach (ObjectId TextStyle_id in Text_style_table)
                        {
                            TextStyleTableRecord TextStyle = (TextStyleTableRecord)Trans1.GetObject(TextStyle_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            if (TextStyle.Name == "Romans_DimStyle")
                            {
                                TextStyle.UpgradeOpen();
                                TextStyle.TextSize = 0;
                                TextStyle.FileName = "romans.shx";
                                TextStyle.TextSize = 0;
                                TextStyle.ObliquingAngle = 0;
                                TextStyle.XScale = 1.0;

                                Text_style_romans = TextStyle;
                                Exista = true;
                                TextStyleID1 = TextStyle.ObjectId;
                                goto Label1;
                            }
                        }


                    Label1:


                        if (Exista == false)
                        {
                            Text_style_table.UpgradeOpen();
                            Text_style_romans = new TextStyleTableRecord();
                            Text_style_romans.Name = "Romans_DimStyle";

                            Text_style_romans.TextSize = 0;
                            Text_style_romans.ObliquingAngle = 0;
                            Text_style_romans.FileName = "romans.shx";
                            Text_style_romans.XScale = 1.0;
                            Text_style_table.Add(Text_style_romans);
                            Trans1.AddNewlyCreatedDBObject(Text_style_romans, true);
                            TextStyleID1 = Text_style_romans.ObjectId;
                        }
                        Trans1.Commit();
                    }

                Start1:
                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the third point (width)");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }


                        Distance_measured = Math.Round((Distance_measured / 5), 0) * 5;
                        Width_measured = Math.Round((Width_measured / 5), 0) * 5;
                        String Continut = Math.Round(Distance_measured, 0).ToString() + "' X " + Math.Round(Width_measured, 0).ToString() + "'";
                        Point3d Position_Mtext = Point_on_poly3;

                        double Rotation1 = 0;

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP4;
                        PP4 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (rotation)");
                        PP4.AllowNone = false;


                        Point_res4 = Editor1.GetPoint(PP4);


                        if (Point_res4.Status == PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                            PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (rotation)");
                            PP5.AllowNone = false;
                            PP5.UseBasePoint = true;
                            PP5.BasePoint = Point_res4.Value;

                            Point_res5 = Editor1.GetPoint(PP5);


                            if (Point_res5.Status == PromptStatus.OK)
                            {
                                Rotation1 = Functions.GET_Bearing_rad(Point_res4.Value.X, Point_res4.Value.Y, Point_res5.Value.X, Point_res5.Value.Y);

                            }

                        }



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_pos;
                        Jig1.jig_Mtext_class Jiggg1;

                        Jiggg1 = new Jig1.jig_Mtext_class(new MText(), 16, Rotation1, Continut, TextStyleID1);
                        Point_pos = Jiggg1.BeginJig();


                        if (Point_pos != null)
                        {
                            Position_Mtext = Point_pos.Value;
                        }







                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Layer = "TEXT";
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = 16;
                        Mtext1.Rotation = Rotation1;
                        Mtext1.TextStyleId = TextStyleID1;
                        Mtext1.Location = Position_Mtext;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);




                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);

                        Trans1.Commit();

                    }



                    goto Start1;

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("ADDFOOT2NUMBER", CommandFlags.UsePickSet)]
        public void ADD_FOOT_TO_NUMBER()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the text or Mtext objects:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;

                                string NEW_TXT1 = Mtext1.Contents;

                                if (Functions.IsNumeric(NEW_TXT1) == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    string NEW_TXT = Mtext1.Contents + "'";
                                    //int ii = 92;
                                    //char cc = (char)ii;
                                    //NEW_TXT = "{" + cc.ToString() + "Fromans|c0;!}";
                                    Mtext1.Contents = NEW_TXT;

                                }




                            }
                            if (Ent1 is DBText)
                            {
                                DBText Text1 = (DBText)Ent1;

                                string NEW_TXT1 = Text1.TextString;

                                if (Functions.IsNumeric(NEW_TXT1) == true)
                                {
                                    Text1.UpgradeOpen();
                                    string NEW_TXT = Text1.TextString + "'";
                                    //int ii = 92;
                                    //char cc = (char)ii;
                                    //NEW_TXT = "{" + cc.ToString() + "Fromans|c0;!}";
                                    Text1.TextString = NEW_TXT;

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
        }

        [CommandMethod("DIM_WIDTH_LEN")]
        public void Measure_regular_ATWS_and_add_it_to_OD_whith_Label_creation_in_no_plot()
        {
            if (isSECURE() == false)
            {
                return;
            }
            string ln = "NO_PLOT_ATWS";

            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {

               

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    Functions.Creaza_layer(ln, 30, false);

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the third point (width)");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }


                        Distance_measured = Math.Round((Distance_measured / 5), 0) * 5;
                        Width_measured = Math.Round((Width_measured / 5), 0) * 5;
                        string Continut = Math.Round(Width_measured, 0).ToString() + "' X " + Math.Round(Distance_measured, 0).ToString() + "'";
                      





                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = 1;
                        Mtext1.Rotation = 0;
                        Mtext1.Location = new Point3d((Punct1.X+Punct3.X)/2, (Punct1.Y + Punct3.Y) / 2,0);
                        Mtext1.Layer = ln;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                        Trans1.Commit();
                        Editor1.WriteMessage(Continut);
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("DIM_WIDTH")]
        public void Measure_width_and_add_it_to_OD_whith_Label_creation_in_no_plot()
        {
            if (isSECURE() == false)
            {
                return;
            }
            string ln = "NO_PLOT_ATWS";

            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {



                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    Functions.Creaza_layer(ln, 30, false);

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (width)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (width)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                       

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);




                        Polyline Poly_width1 = new Polyline();


                        Poly_width1.AddVertexAt(0, new Point2d(Point_on_poly1.X, Point_on_poly1.Y), 0, 0, 0);
                       
                        Poly_width1.AddVertexAt(1, new Point2d(Point_on_poly2.X, Point_on_poly2.Y), 0, 0, 0);

                        double Width_measured = Poly_width1.Length;

                        Width_measured = Functions.Round5(Width_measured);

                        string Continut = Functions.Get_String_Rounded(Width_measured,0) ;






                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = 1;
                        Mtext1.Rotation = 0;
                        Mtext1.Location = new Point3d((Punct1.X + Punct2.X) / 2, (Punct1.Y + Punct2.Y) / 2, 0);
                        Mtext1.Layer = ln;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                        Trans1.Commit();
                        Editor1.WriteMessage(Continut);
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("DIM_ATWS_R")]
        public void Measure_regular_ATWS_and_add_it_to_OD()
        {
            if (isSECURE() == false)
            {
                return;
            }


            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {

                int TextHeight1 = 8;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the third point (width)");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }


                        Distance_measured = Math.Round((Distance_measured / 5), 0) * 5;
                        Width_measured = Math.Round((Width_measured / 5), 0) * 5;
                        String Continut = Math.Round(Distance_measured, 0).ToString() + "' X " + Math.Round(Width_measured, 0).ToString() + "'";
                        Point3d Position_Mtext = Point_on_poly3;

                        double Rotation1 = 0;



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_pos;
                        Jig1.jig_Mtext_class Jiggg1;

                        Jiggg1 = new Jig1.jig_Mtext_class(new MText(), TextHeight1 * 2, Rotation1, Continut, TextStyleID1);
                        Point_pos = Jiggg1.BeginJig();


                        if (Point_pos != null)
                        {
                            Position_Mtext = Point_pos.Value;
                        }


                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = TextHeight1;
                        Mtext1.Rotation = Rotation1;
                        Mtext1.TextStyleId = TextStyleID1;
                        Mtext1.Location = Position_Mtext;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);




                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        Trans1.Commit();
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("DIM_ATWS_N")]
        public void Measure_irregular_ATWS_and_add_it_to_OD()
        {
            if (isSECURE() == false)
            {
                return;
            }


            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;


                Autodesk.AutoCAD.EditorInput.PromptIntegerOptions Prompt_int = new Autodesk.AutoCAD.EditorInput.PromptIntegerOptions("\n" + "Specify rounding:");
                Prompt_int.AllowNegative = false;
                Prompt_int.AllowZero = true;
                Prompt_int.AllowNone = true;
                Prompt_int.DefaultValue = 0;
                Prompt_int.UseDefaultValue = true;

                Autodesk.AutoCAD.EditorInput.PromptIntegerResult Rezultat_rounding = ThisDrawing.Editor.GetInteger(Prompt_int);

                int TextHeight1 = 8;




                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the other side of the polyline");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }

                        int Rounding = Rezultat_rounding.Value;
                        if (Rounding == 0) Rounding = 1;

                        Distance_measured = Math.Round((Distance_measured / Rounding), 0) * Rounding;


                        Width_measured = Math.Round(((Poly1.Area / Distance_measured) / Rounding), 0) * Rounding;


                        String Continut = Math.Round(Distance_measured, 0).ToString() + "' X " + Math.Round(Width_measured, 0).ToString() + "'*";
                        Point3d Position_Mtext = Point_on_poly3;

                        double Rotation1 = 0;



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_pos;
                        Jig1.jig_Mtext_class Jiggg1;

                        Jiggg1 = new Jig1.jig_Mtext_class(new MText(), TextHeight1 * 2, Rotation1, Continut, TextStyleID1);
                        Point_pos = Jiggg1.BeginJig();


                        if (Point_pos != null)
                        {
                            Position_Mtext = Point_pos.Value;
                        }


                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = TextHeight1;
                        Mtext1.Rotation = Rotation1;
                        Mtext1.TextStyleId = TextStyleID1;
                        Mtext1.Location = Position_Mtext;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);




                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        Trans1.Commit();
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("ATWS1")]
        public void Measure_regular_ATWS_and_add_it_to_OD_SPIRE()
        {
            if (isSECURE() == false)
            {
                return;
            }


            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {

                int TextHeight1 = 8;
                int Rounding = 5;

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the third point (width)");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }


                        Distance_measured = Math.Round((Distance_measured / Rounding), 0) * Rounding;
                        Width_measured = Math.Round((Width_measured / Rounding), 0) * Rounding;
                        String Continut = Math.Round(Distance_measured, 0).ToString() + "' X " + Math.Round(Width_measured, 0).ToString() + "'";
                        Point3d Position_Mtext = Point_on_poly3;

                        double Rotation1 = 0;



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_pos;
                        Jig1.jig_Mtext_class Jiggg1;

                        Jiggg1 = new Jig1.jig_Mtext_class(new MText(), TextHeight1 * 2, Rotation1, Continut, TextStyleID1);
                        Point_pos = Jiggg1.BeginJig();


                        if (Point_pos != null)
                        {
                            Position_Mtext = Point_pos.Value;
                        }


                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = TextHeight1;
                        Mtext1.Rotation = Rotation1;
                        Mtext1.TextStyleId = TextStyleID1;
                        Mtext1.Location = Position_Mtext;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                        Poly1.UpgradeOpen();
                        Poly1.ColorIndex = 1;


                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        Trans1.Commit();
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        [CommandMethod("ATWS2")]
        public void Measure_irregular_ATWS_and_add_it_to_OD_SPIRE()
        {
            if (isSECURE() == false)
            {
                return;
            }


            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;



                int TextHeight1 = 8;
                int Rounding = 5;



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                Start1:

                    ObjectId ObjID1 = ObjectId.Null;
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        object OLD_OSnap = Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("OSMODE");
                        object NEW_OSnap = 512;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_atws;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_atws;
                        Prompt_atws = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ATWS:");
                        Prompt_atws.SetRejectMessage("\nSelect a polyline!");
                        Prompt_atws.AllowNone = true;
                        Prompt_atws.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_atws = ThisDrawing.Editor.GetEntity(Prompt_atws);

                        if (Rezultat_atws.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_atws.ObjectId, OpenMode.ForRead);

                        //End = 1,
                        //Middle = 2,
                        //Center = 4,
                        //Node = 8,
                        //Quadrant = 16,
                        //Intersection = 32,
                        //Insertion = 64,
                        //Perpendicular = 128,
                        //Tangent = 256,
                        //Near = 512,
                        // Quick = 1024,
                        //ApparentIntersection = 2048,
                        //Immediate = 65536,
                        //AllowTangent = 131072,
                        // DisablePerpendicular = 262144,
                        //RelativeCartesian = 524288,
                        //RelativePolar = 1048576,
                        //NoneOverride = 2097152,                

                        NEW_OSnap = 545;

                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", NEW_OSnap);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point (length)");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point (length)");
                        PP2.AllowNone = false;
                        PP2.UseBasePoint = true;
                        PP2.BasePoint = Point_res1.Value;

                        Point_res2 = Editor1.GetPoint(PP2);


                        if (Point_res2.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                        PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the other side of the polyline");
                        PP3.AllowNone = false;
                        PP3.UseBasePoint = true;
                        PP3.BasePoint = Point_res2.Value;

                        Point_res3 = Editor1.GetPoint(PP3);


                        if (Point_res3.Status != PromptStatus.OK)
                        {

                            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);
                            Trans1.Commit();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Point3d Punct1 = new Point3d();
                        Punct1 = Point_res1.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct2 = new Point3d();
                        Punct2 = Point_res2.Value.TransformBy(CurentUCSmatrix);
                        Point3d Punct3 = new Point3d();
                        Punct3 = Point_res3.Value.TransformBy(CurentUCSmatrix);

                        Point3d Point_on_poly1 = new Point3d();
                        Point_on_poly1 = Poly1.GetClosestPointTo(Punct1, Vector3d.ZAxis, false);
                        Point3d Point_on_poly2 = new Point3d();
                        Point_on_poly2 = Poly1.GetClosestPointTo(Punct2, Vector3d.ZAxis, false);
                        Point3d Point_on_poly3 = new Point3d();
                        Point_on_poly3 = Poly1.GetClosestPointTo(Punct3, Vector3d.ZAxis, false);

                        Double Param1 = Poly1.GetParameterAtPoint(Point_on_poly1);
                        Double Param2 = Poly1.GetParameterAtPoint(Point_on_poly2);


                        if (Param1 > Param2)
                        {
                            double T = Param1;
                            Param1 = Param2;
                            Param2 = T;

                            Point3d tp = new Point3d();
                            tp = Point_on_poly1;
                            Point_on_poly1 = Point_on_poly2;
                            Point_on_poly2 = tp;
                        }

                        Polyline Poly_length1 = new Polyline();
                        int Index_poly1 = 0;

                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly1 = Index_poly1 + 1;
                        if (Math.Floor(Param2) - Math.Floor(Param1) >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Ceiling(Param1)); i <= Math.Floor(Param2); i = i + 1)
                            {
                                Poly_length1.AddVertexAt(Index_poly1, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly1 = Index_poly1 + 1;
                            }
                        }
                        Poly_length1.AddVertexAt(Index_poly1, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        Polyline Poly_length2 = new Polyline();
                        int Index_poly2 = 0;
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param1).X, Poly1.GetPointAtParameter(Param1).Y), 0, 0, 0);
                        Index_poly2 = Index_poly2 + 1;
                        if (Param1 >= 1)
                        {
                            for (int i = Convert.ToInt32(Math.Floor(Param1)); i >= 0; i = i - 1)
                            {
                                Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                                Index_poly2 = Index_poly2 + 1;
                            }
                        }

                        for (int i = Poly1.NumberOfVertices - 1; i >= Convert.ToInt32(Math.Ceiling(Param2)); i = i - 1)
                        {
                            Poly_length2.AddVertexAt(Index_poly2, Poly1.GetPoint2dAt(i), 0, 0, 0);
                            Index_poly2 = Index_poly2 + 1;
                        }
                        Poly_length2.AddVertexAt(Index_poly2, new Point2d(Poly1.GetPointAtParameter(Param2).X, Poly1.GetPointAtParameter(Param2).Y), 0, 0, 0);

                        double Distance_measured = Poly_length1.Length;
                        double Width_measured = Poly_length1.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;

                        if (Width_measured == 0)
                        {
                            Distance_measured = Poly_length2.Length;
                            Width_measured = Poly_length2.GetClosestPointTo(Point_on_poly3, Vector3d.ZAxis, true).GetVectorTo(Point_on_poly3).Length;
                        }




                        Distance_measured = Math.Round((Distance_measured / Rounding), 0) * Rounding;


                        Width_measured = Math.Round(((Poly1.Area / Distance_measured) / Rounding), 0) * Rounding;


                        String Continut = Math.Round(Distance_measured, 0).ToString() + "' X " + Math.Round(Width_measured, 0).ToString() + "'";
                        Point3d Position_Mtext = Point_on_poly3;

                        double Rotation1 = 0;



                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_pos;
                        Jig1.jig_Mtext_class Jiggg1;

                        Jiggg1 = new Jig1.jig_Mtext_class(new MText(), TextHeight1 * 2, Rotation1, Continut, TextStyleID1);
                        Point_pos = Jiggg1.BeginJig();


                        if (Point_pos != null)
                        {
                            Position_Mtext = Point_pos.Value;
                        }


                        MText Mtext1 = new MText();
                        Mtext1.Contents = Continut;
                        Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                        Mtext1.TextHeight = TextHeight1;
                        Mtext1.Rotation = Rotation1;
                        Mtext1.TextStyleId = TextStyleID1;
                        Mtext1.Location = Position_Mtext;
                        BTrecord.AppendEntity(Mtext1);
                        Trans1.AddNewlyCreatedDBObject(Mtext1, true);

                        Poly1.UpgradeOpen();
                        Poly1.ColorIndex = 1;


                        Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSMODE", OLD_OSnap);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_atws.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                        {
                            if (Records1 != null)
                            {

                                if (Records1.Count > 0)
                                {
                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                    {
                                        Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                        for (int i = 0; i < Record1.Count; ++i)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                            string Nume_field = Field_def1.Name;
                                            if (Nume_field.ToUpper() == "NOTE1")
                                            {
                                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                                Valoare1 = Record1[i];
                                                Valoare1.Assign(Continut);
                                                Records1.UpdateRecord(Record1);
                                                i = Record1.Count;
                                            }

                                        }
                                    }
                                }
                            }

                        }

                        Trans1.Commit();
                    }

                    goto Start1;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }




        [CommandMethod("CHAINAGE2STATION")]
        public void STATIONS_CANADA_US()
        {
            if (isSECURE() == false)
            {
                return;
            }

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the text or Mtext objects:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is DBText)
                            {
                                DBText Text1 = (DBText)Ent1;
                                if (Functions.IsNumeric(Text1.TextString.Replace("+", "")) == true)
                                {
                                    Text1.UpgradeOpen();
                                    double Station = Convert.ToDouble(Text1.TextString.Replace("+", ""));
                                    Text1.TextString = Functions.Get_chainage_feet_from_double(Station, 0);
                                }
                            }
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;
                                if (Functions.IsNumeric(Mtext1.Contents.Replace("+", "")) == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    double Station = Convert.ToDouble(Mtext1.Contents.Replace("+", ""));
                                    Mtext1.Contents = Functions.Get_chainage_feet_from_double(Station, 0);
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
        }

        [CommandMethod("feet2rod")]
        public void CONVERT_FEET_TO_RODS()
        {
            if (isSECURE() == false)
            {
                return;
            }

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is DBText)
                            {
                                DBText Text1 = (DBText)Ent1;
                                if (Text1.TextString.Contains("'") == true)
                                {
                                    if (Functions.IsNumeric(Text1.TextString.Replace("'", "")) == true)
                                    {
                                        Text1.UpgradeOpen();
                                        double Foot = Convert.ToDouble(Text1.TextString.Replace("'", ""));
                                        Text1.TextString = Functions.Get_String_Rounded(Foot / 16.5, 2);
                                    }
                                }
                            }
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;
                                if (Mtext1.Contents.Contains("'") == true)
                                {
                                    if (Functions.IsNumeric(Mtext1.Contents.Replace("'", "")) == true)
                                    {
                                        Mtext1.UpgradeOpen();
                                        double Foot = Convert.ToDouble(Mtext1.Contents.Replace("'", ""));
                                        Mtext1.Contents = Functions.Get_String_Rounded(Foot / 16.5, 2);
                                    }
                                }
                            }

                            if (Ent1 is BlockReference)
                            {
                                BlockReference Block1 = (BlockReference)Ent1;
                                if (Block1.AttributeCollection.Count > 0)
                                {
                                    Block1.UpgradeOpen();
                                    foreach (ObjectId Id in Block1.AttributeCollection)
                                    {
                                        if (Id.IsErased == false)
                                        {
                                            AttributeReference attRef = (AttributeReference)Trans1.GetObject(Id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                            String Continut = attRef.TextString;
                                            if (Continut.Contains("'") == true)
                                            {
                                                if (Functions.IsNumeric(Continut.Replace("'", "")) == true)
                                                {

                                                    double Foot = Convert.ToDouble(Continut.Replace("'", ""));
                                                    attRef.TextString = Functions.Get_String_Rounded(Foot / 16.5, 2);
                                                }
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
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("add_now_or_F")]
        public void add_now_or_formerly()
        {
            if (isSECURE() == false)
            {
                return;
            }

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is DBText)
                            {
                                DBText Text1 = (DBText)Ent1;
                                if (Text1.TextString.Contains("N/F ") == false)
                                {
                                    Text1.UpgradeOpen();

                                    Text1.TextString = "N/F " + Text1.TextString;
                                }

                            }
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;

                                if (Mtext1.Contents.Contains("N/F ") == false)
                                {
                                    Mtext1.UpgradeOpen();

                                    Mtext1.Contents = "N/F " + Mtext1.Contents;
                                }


                            }

                            if (Ent1 is BlockReference)
                            {
                                BlockReference Block1 = (BlockReference)Ent1;
                                if (Block1.AttributeCollection.Count > 0)
                                {
                                    Block1.UpgradeOpen();
                                    foreach (ObjectId Id in Block1.AttributeCollection)
                                    {
                                        if (Id.IsErased == false)
                                        {
                                            AttributeReference attRef = (AttributeReference)Trans1.GetObject(Id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                                            if (attRef.Tag.ToUpper() == "OWNER" | attRef.Tag.ToUpper() == "PARCEL_OWNER")
                                            {

                                                if (attRef.TextString.Contains("N/F ") == false)
                                                {
                                                    attRef.TextString = "N/F " + attRef.TextString;
                                                }

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
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        //[CommandMethod("MAGICNUMBER")]
        public void MAGIC_NUMBER()
        {
            if (isSECURE() == false)
            {
                return;
            }
            int Magic_number = 130;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the objects:";
                        Prompt_rez.SingleOnly = true;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);

                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;
                                string String_de_procesat = Mtext1.Contents;
                                Mtext1.UpgradeOpen();
                                Mtext1.Width = Magic_number;



                                if (String_de_procesat.Length > Magic_number)
                                {
                                    if (String_de_procesat.Contains(" ") == true)
                                    {
                                        string Noul_string;
                                        int Pozitie_space1;

                                        Pozitie_space1 = String_de_procesat.IndexOf(" ");
                                        //MessageBox.Show(Pozitie_space1.ToString());
                                        //if (Pozitie_space1 < Magic_number)
                                        //{
                                        //int (pos1
                                        //}


                                        //string String_rest;
                                        //String_rest = String_de_procesat.Substring( Pozitie_space1 + 1);
                                        //MessageBox.Show(String_rest);


                                        if (Pozitie_space1 + 1 > Magic_number)
                                        {

                                        }


                                    }



                                }




                            }

                            if (Ent1 is BlockReference)
                            {
                                BlockReference Block1 = (BlockReference)Ent1;
                                if (Block1.AttributeCollection.Count > 0)
                                {
                                    Block1.UpgradeOpen();
                                    foreach (ObjectId Id in Block1.AttributeCollection)
                                    {
                                        if (Id.IsErased == false)
                                        {
                                            AttributeReference attRef = (AttributeReference)Trans1.GetObject(Id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                            attRef.MTextAttribute.Width = Magic_number;

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










        }

        [CommandMethod("REPLACE_angle", CommandFlags.UsePickSet)]
        public void REPLACE_ANGLE_WITH_ANGLE()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the Mtext objects:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;
                                string NEW_TXT1 = Mtext1.Contents;
                                if (Mtext1.Contents.Contains("<") == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    string Continut = Mtext1.Contents;
                                    Continut = Continut.Replace("\\L", "");
                                    Continut = Continut.Replace("<", "{\\f@Arial Unicode MS|b0|i0|c0|p34;\\H22;∡}");
                                    Mtext1.Contents = Continut;
                                    Mtext1.Location = new Point3d(Mtext1.Location.X - 3, Mtext1.Location.Y, 0);
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
        }
        [CommandMethod("ZZ1", CommandFlags.UsePickSet)]
        public void ZZ()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the Mtext objects:";
                            Prompt_rez.SingleOnly = true;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;

                                string NEW_TXT1 = Mtext1.Contents;
                                System.Windows.Forms.Clipboard.SetText(NEW_TXT1);
                                MessageBox.Show(NEW_TXT1);

                                if (Mtext1.Contents.Contains("<") == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    //string NEW_TXT = "{\fromans|c0;!}";
                                    //int ii = 92;
                                    //char cc = (char)ii;
                                    //NEW_TXT = "{" + cc.ToString() + "Fromans|c0;!}";
                                    string Continut = Mtext1.Contents;
                                    Mtext1.Contents = Continut.Replace("<", "{\f@Arial Unicode MS|b0|p34;xx}");

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
        }


        [CommandMethod("dim_q_bearing", CommandFlags.UsePickSet)]
        public void dim_q_bearing()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is Polyline)
                            {
                                double TextH = 16;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1;

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.MiddleCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




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
        }


        [CommandMethod("dim_BD0", CommandFlags.UsePickSet)]
        public void dim_BD0()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is Polyline)
                            {
                                double TextH = 16;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 0);

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




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
        }


        [CommandMethod("dim_BD1", CommandFlags.UsePickSet)]
        public void dim_BD1()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is Polyline)
                            {
                                double TextH = 16;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 1);

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




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
        }

        [CommandMethod("dim_BD2", CommandFlags.UsePickSet)]
        public void dim_BD2()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is Polyline)
                            {
                                double TextH = 16;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 2);

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




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
        }



        [CommandMethod("dim_BD3", CommandFlags.UsePickSet)]
        public void dim_BD3()
        {
            if (isSECURE() == false)
            {
                return;
            }

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is Polyline)
                            {
                                double TextH = 16;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }

                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));
                                    }

                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);
                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 3);
                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);
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
        }

        [CommandMethod("dim_BD3D0", CommandFlags.UsePickSet)]
        public void dim_BD3D0()
        {
            if (isSECURE() == false)
            {
                return;
            }


            string vbcrlf = System.Environment.NewLine;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the polyline:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        String Dist_val = "";

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);


                            if (Ent1 is Polyline3d)
                            {
                                double TextH = 1;
                                Polyline3d Poly3d = (Polyline3d)Ent1;
                                Polyline Poly1 = new Polyline();
                                int Index2d1 = 0;

                                foreach (ObjectId ObjId in Poly3d)
                                {
                                    PolylineVertex3d vertex1 = (PolylineVertex3d)Trans1.GetObject(ObjId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                    if (vertex1 != null)
                                    {
                                        Poly1.AddVertexAt(Index2d1, new Point2d(vertex1.Position.X, vertex1.Position.Y), 0, 0, 0);
                                        Index2d1 = Index2d1 + 1;
                                    }
                                }


                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = new Point3d(Pt1.X, Pt1.Y, Poly3d.GetPointAtParameter(j).Z).GetVectorTo(new Point3d(Pt2.X, Pt2.Y, Poly3d.GetPointAtParameter(j + 1).Z)).Length;
                                    Dist_val = Dist_val + vbcrlf + Dist1.ToString();
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 0);

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly3d.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




                                }





                            }


                            if (Ent1 is Polyline)
                            {
                                double TextH = 1;
                                Polyline Poly1 = (Polyline)Ent1;
                                for (int j = 0; j < Poly1.NumberOfVertices - 1; ++j)
                                {
                                    double x1 = Poly1.GetPoint3dAt(j).X;
                                    double y1 = Poly1.GetPoint3dAt(j).Y;
                                    double x2 = Poly1.GetPoint3dAt(j + 1).X;
                                    double y2 = Poly1.GetPoint3dAt(j + 1).Y;
                                    Point3d Pt1 = Poly1.GetPoint3dAt(j);
                                    Point3d Pt2 = Poly1.GetPoint3dAt(j + 1);
                                    double Dist1 = Pt1.GetVectorTo(Pt2).Length;
                                    Dist_val = Dist_val + vbcrlf + Dist1.ToString();
                                    double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);

                                    string Prefix1 = "N";
                                    string Suffix1 = "E";
                                    Double Quadrant1 = Math.PI / 2 - Bearing1;

                                    if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                                    {
                                        Quadrant1 = Bearing1 - Math.PI / 2;
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                        Prefix1 = "S";
                                        Suffix1 = "W";
                                    }

                                    if (Bearing1 > 3 * Math.PI / 2)
                                    {
                                        Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                        Prefix1 = "S";
                                        Suffix1 = "E";
                                    }


                                    Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                                    Line Linie1 = new Line(PointM, Pt2);
                                    Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                                    Point3d Point_ins = new Point3d();
                                    if (Linie1.Length < TextH)
                                    {
                                        Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                                    }
                                    Point_ins = Linie1.GetPointAtDist(TextH / 2);

                                    string Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 3);

                                    MText Mtext1 = new MText();
                                    Mtext1.Contents = Content1;
                                    Mtext1.Layer = Poly1.Layer;
                                    Mtext1.Rotation = Bearing1;
                                    Mtext1.TextHeight = TextH;
                                    Mtext1.Location = Point_ins;
                                    Mtext1.Attachment = AttachmentPoint.BottomCenter;
                                    BTrecord.AppendEntity(Mtext1);
                                    Trans1.AddNewlyCreatedDBObject(Mtext1, true);




                                }


                                System.Windows.Forms.Clipboard.SetText(Dist_val);


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


        [CommandMethod("dim_dist0")]
        public void dim_segment()
        {
            if (isSECURE() == false)
            {
                return;
            }


            string vbcrlf = System.Environment.NewLine;

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the segment:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline or a line!\nCommon man! Pay attention man!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Line), false);
                        Rezultat1 = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }



                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        Entity Ent1 = (Entity)Trans1.GetObject(Rezultat1.ObjectId, OpenMode.ForRead);

                        Point3d Pickpt = Rezultat1.PickedPoint;
                        double TextH = 18;


                        if (Ent1 is Polyline)
                        {
                            Polyline Poly1 = (Polyline)Ent1;


                            double Param1 = Poly1.GetParameterAtPoint(Poly1.GetClosestPointTo(Pickpt, Vector3d.ZAxis, false));
                            Point3d Pt1 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Floor(Param1)));
                            Point3d Pt2 = Poly1.GetPoint3dAt(Convert.ToInt32(Math.Ceiling(Param1)));

                            double x1 = Pt1.X;
                            double y1 = Pt1.Y;
                            double x2 = Pt2.X;
                            double y2 = Pt2.Y;
                            double Dist1 = Math.Pow((Math.Pow((x1 - x2), 2) + Math.Pow((y1 - y2), 2)), 0.5);
                            double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);


                            Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                            Line Linie1 = new Line(PointM, Pt2);
                            Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                            Point3d Point_ins = new Point3d();
                            if (Linie1.Length < TextH)
                            {
                                Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                            }
                            Point_ins = Linie1.GetPointAtDist(TextH / 2);

                            string Content1 = Functions.Get_String_Rounded(Dist1, 0) + "'";

                            MText Mtext1 = new MText();
                            Mtext1.Contents = Content1;
                            Mtext1.Rotation = Bearing1;
                            Mtext1.TextHeight = TextH;
                            Mtext1.Location = Point_ins;
                            Mtext1.Attachment = AttachmentPoint.BottomCenter;
                            BTrecord.AppendEntity(Mtext1);
                            Trans1.AddNewlyCreatedDBObject(Mtext1, true);
                        }

                        if (Ent1 is Line)
                        {
                            Line Line1 = (Line)Ent1;


                            Point3d Pt1 = Line1.StartPoint;
                            Point3d Pt2 = Line1.EndPoint;

                            double x1 = Pt1.X;
                            double y1 = Pt1.Y;
                            double x2 = Pt2.X;
                            double y2 = Pt2.Y;
                            double Dist1 = Math.Pow((Math.Pow((x1 - x2), 2) + Math.Pow((y1 - y2), 2)), 0.5);
                            double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);


                            Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                            Line Linie1 = new Line(PointM, Pt2);
                            Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                            Point3d Point_ins = new Point3d();
                            if (Linie1.Length < TextH)
                            {
                                Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));

                            }
                            Point_ins = Linie1.GetPointAtDist(TextH / 2);


                            string Content1 = Functions.Get_String_Rounded(Dist1, 0) + "'";

                            MText Mtext1 = new MText();
                            Mtext1.Contents = Content1;

                            Mtext1.Rotation = Bearing1;
                            Mtext1.TextHeight = TextH;
                            Mtext1.Location = Point_ins;
                            Mtext1.Attachment = AttachmentPoint.BottomCenter;
                            BTrecord.AppendEntity(Mtext1);
                            Trans1.AddNewlyCreatedDBObject(Mtext1, true);

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

        [CommandMethod("CLEAN_STREAM_WETLAND", CommandFlags.UsePickSet)]
        public void REPLACE_in_Mtext_STREAM_WETLAND()
        {
            if (isSECURE() == false)
            {
                return;
            }


            try
            {

                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                            Rezultat1 = Editor1.SelectImplied();

                            if (Rezultat1.Status != PromptStatus.OK)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the text or Mtext objects:";
                                Prompt_rez.SingleOnly = true;
                                Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                            }



                            if (Rezultat1.Status != PromptStatus.OK)
                            {

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                                Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                                if (Ent1 is MText)
                                {
                                    MText Mtext1 = (MText)Ent1;



                                    if (Mtext1.Text.Contains("STREAM") == true)
                                    {
                                        Mtext1.UpgradeOpen();
                                        Mtext1.Contents = "STREAM";

                                    }
                                    if (Mtext1.Text.Contains("WETLAND") == true)
                                    {
                                        Mtext1.UpgradeOpen();
                                        Mtext1.Contents = "WETLAND";

                                    }



                                }

                                if (Ent1 is MLeader)
                                {

                                    MLeader Mleader1 = (MLeader)Ent1;

                                    MText Mtext1 = Mleader1.MText;



                                    if (Mtext1.Text.Contains("STREAM") == true)
                                    {
                                        Mleader1.UpgradeOpen();
                                        Mtext1.Contents = "STREAM";
                                        Mleader1.MText = Mtext1;
                                    }
                                    if (Mtext1.Text.Contains("WETLAND") == true)
                                    {
                                        Mleader1.UpgradeOpen();
                                        Mtext1.Contents = "WETLAND";
                                        Mleader1.MText = Mtext1;
                                    }



                                }


                                if (Ent1 is DBText)
                                {

                                    DBText Text1 = (DBText)Ent1;





                                    if (Text1.TextString.Contains("STREAM") == true)
                                    {
                                        Text1.UpgradeOpen();
                                        Text1.TextString = "STREAM";

                                    }
                                    if (Text1.TextString.Contains("WETLAND") == true)
                                    {
                                        Text1.UpgradeOpen();
                                        Text1.TextString = "WETLAND";

                                    }



                                }

                            }






                            Trans1.Commit();
                            MessageBox.Show("Done");
                        }



                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        [CommandMethod("underline_defl", CommandFlags.UsePickSet)]
        public void add_line_to_defl()
        {
            if (isSECURE() == false)
            {
                return;
            }




            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Rezultat1 = Editor1.SelectImplied();

                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                            Prompt_rez.MessageForAdding = "\nSelect the Mtext objects:";
                            Prompt_rez.SingleOnly = false;
                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        }



                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            Autodesk.AutoCAD.EditorInput.SelectedObject Obj1 = Rezultat1.Value[i];
                            Entity Ent1 = (Entity)Obj1.ObjectId.GetObject(OpenMode.ForRead);
                            if (Ent1 is MText)
                            {
                                MText Mtext1 = (MText)Ent1;




                                int ii = 92;
                                char cc = (char)ii;
                                string Slash = cc.ToString();

                                string Continut = Mtext1.Contents;

                                if (Continut.Contains("{" + Slash + "W0.8;" + Slash + "L") == true)
                                {
                                    Mtext1.UpgradeOpen();
                                    Mtext1.Contents = Continut.Replace("{" + Slash + "W0.8;" + Slash + "L", "{" + Slash + "W0.8;");
                                    Line Line1 = new Line(new Point3d(Mtext1.Location.X + 3, Mtext1.Location.Y - 75, 0), new Point3d(Mtext1.Location.X + 3, Mtext1.Location.Y + 625, 0));
                                    Line1.Layer = Mtext1.Layer;
                                    BTrecord.AppendEntity(Line1);
                                    Trans1.AddNewlyCreatedDBObject(Line1, true);

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
        }

        [CommandMethod("dim_PICKPT")]
        public void dim_PICK_POINT()
        {
            if (isSECURE() == false)
            {
                return;
            }

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {


                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_poly1;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_Poly1;
                        Prompt_Poly1 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the polyline:");
                        Prompt_Poly1.SetRejectMessage("\nSelect a line, polyline or arc!");
                        Prompt_Poly1.AllowNone = true;
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Line), false);
                        Prompt_Poly1.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Arc), false);

                        Rezultat_poly1 = ThisDrawing.Editor.GetEntity(Prompt_Poly1);

                        if (Rezultat_poly1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Point3d Point1 = new Point3d();
                        Point1 = Point_res1.Value;

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                        PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point");
                        PP2.AllowNone = false;
                        PP2.BasePoint = Point1;
                        PP2.UseBasePoint = true;
                        Point_res2 = Editor1.GetPoint(PP2);

                        if (Point_res2.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }
                        Point3d Point2 = new Point3d();
                        Point2 = Point_res2.Value;


                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_poly1.ObjectId, OpenMode.ForRead);
                        if (Ent1 is Polyline | Ent1 is Autodesk.AutoCAD.DatabaseServices.Line)
                        {
                            double TextH = 16;
                            Curve Curve1 = (Curve)Ent1;

                            Point3d Pt1 = new Point3d();
                            Point3d Pt2 = new Point3d();
                            Pt1 = Curve1.GetClosestPointTo(Point1, Vector3d.ZAxis, false);
                            Pt2 = Curve1.GetClosestPointTo(Point2, Vector3d.ZAxis, false);

                            double x1 = Pt1.X;
                            double y1 = Pt1.Y;
                            double x2 = Pt2.X;
                            double y2 = Pt2.Y;

                            double Dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);

                            double Bulge1 = 0;
                            double Radius1 = 0;
                            double Unghi_la_centru = 0;

                            string Content1 = "";

                            if (Ent1 is Polyline)
                            {
                                Polyline Curve11 = (Polyline)Ent1;
                                double Param1 = Curve11.GetParameterAtPoint(Pt1);
                                double Param2 = Curve11.GetParameterAtPoint(Pt2);
                                if (Math.Round(Param1, 4) >= Math.Round(Param2, 4))
                                {
                                    double T = Param1;
                                    Param1 = Param2;
                                    Param2 = T;
                                }

                                if (Math.Round(Param2, 4) - Math.Round(Param1, 4) <= 1)
                                {
                                    double Arc_dist = Curve11.GetDistanceAtParameter(Param2) - Curve11.GetDistanceAtParameter(Param1);

                                    Bulge1 = Curve11.GetBulgeAt(Convert.ToInt32(Math.Floor(Param1)));
                                    Bulge1 = Math.Abs(4 * Math.Atan(Bulge1) * 180 / Math.PI);

                                    double d = Curve11.GetPointAtParameter(Math.Floor(Param1)).GetVectorTo(Curve11.GetPointAtParameter(Math.Floor(Param1) + 1)).Length / 2;
                                    double b = 0.5 * Bulge1 * Math.PI / 180;

                                    Radius1 = d / Math.Sin(b);

                                    Unghi_la_centru = (2 * Math.Asin((Dist1 / 2) / Radius1)) * 180 / Math.PI;

                                    Content1 = "R = " + Math.Round(Radius1, 4).ToString() + "'\\PArc dist = " + Math.Round(Arc_dist, 4).ToString() + "'\\PDelta = " + Math.Round(Unghi_la_centru, 4).ToString();

                                }
                            }

                            double Bearing1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                            string Prefix1 = "N";
                            string Suffix1 = "E";
                            Double Quadrant1 = Math.PI / 2 - Bearing1;



                            if (Bearing1 > Math.PI / 2 & Bearing1 <= Math.PI)
                            {
                                Quadrant1 = Bearing1 - Math.PI / 2;
                                Suffix1 = "W";
                            }

                            if (Bearing1 > Math.PI & Bearing1 <= 3 * Math.PI / 2)
                            {
                                Quadrant1 = 3 * Math.PI / 2 - Bearing1;
                                Prefix1 = "S";
                                Suffix1 = "W";
                            }

                            if (Bearing1 > 3 * Math.PI / 2)
                            {
                                Quadrant1 = Bearing1 - 3 * Math.PI / 2;
                                Prefix1 = "S";
                                Suffix1 = "E";
                            }

                            Point3d PointM = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                            Line Linie1 = new Line(PointM, new Point3d(x2, y2, 0));
                            Linie1.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, PointM));
                            Point3d Point_ins = new Point3d();
                            if (Linie1.Length < TextH)
                            {
                                Linie1.TransformBy(Matrix3d.Scaling(1.1 * TextH / Linie1.Length, PointM));
                            }

                            Point_ins = Linie1.GetPointAtDist(TextH / 2);
                            if (Bulge1 == 0)
                            {
                                Content1 = Prefix1 + Functions.Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1 + "\\P" + Functions.Get_String_Rounded(Dist1, 4);
                            }
                            MText Mtext1 = new MText();
                            Mtext1.Contents = Content1;
                            Mtext1.Layer = Curve1.Layer;
                            Mtext1.Rotation = Bearing1;
                            Mtext1.TextHeight = TextH;
                            Mtext1.Location = Point_ins;
                            Mtext1.Attachment = AttachmentPoint.BottomCenter;
                            BTrecord.AppendEntity(Mtext1);
                            Trans1.AddNewlyCreatedDBObject(Mtext1, true);

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



        [CommandMethod("BRD")]
        public void Show_BEARING_AND_DIST_FORM()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Bearing_and_dist_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    Bearing_and_dist_form forma2 = new Bearing_and_dist_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }

        [CommandMethod("vv")]
        public void Test1()
        {
            if (isSECURE() == true)
            {
                System.Collections.Specialized.StringCollection col1 = new System.Collections.Specialized.StringCollection();
                System.Collections.Specialized.StringCollection col2 = new System.Collections.Specialized.StringCollection();
                col1.Add("CURVE");
                col2.Add("123");
                col1.Add("DELTA");
                col2.Add("");
                col1.Add("RADIUS");
                col2.Add("456");
                Functions.InsertBlock_with_multiple_atributes("C:\\Users\\pop70694\\Documents\\BLOCKS\\Curve_Table.dwg", "Curve_Table", new Point3d(0, 0, 0), 1, "0", col1, col2);
                Functions.InsertBlock_with_multiple_atributes("C:\\Users\\pop70694\\Documents\\BLOCKS\\Line_Table.dwg", "Line_Table", new Point3d(0, 0, 0), 1, "0", col1, col2);
                Functions.InsertBlock_with_multiple_atributes("C:\\Users\\pop70694\\Documents\\BLOCKS\\PI_BLOCK.dwg", "PI_BLOCK", new Point3d(0, 0, 0), 1, "0", col1, col2);
            }
            else
            {
                return;
            }
        }


        [CommandMethod("RVA")]
        public void REMOVE_VERTEX_FROM_POLY()
        {
            if (isSECURE() == false)
            {
                return;
            }

            //arrowhead parameters

            // text parameters

            ObjectId DimStyleID1 = ObjectId.Null;
            ObjectId TextStyleID1 = ObjectId.Null;
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;



                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the polyline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }

                        Polyline Poly1 = (Polyline)Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForWrite);

                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_result1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_res1;
                        PP_res1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the first point");
                        PP_res1.AllowNone = false;
                        Point_result1 = Editor1.GetPoint(PP_res1);

                        if (Point_result1.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_result2;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_res2;
                        PP_res2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the second point\r\n(side)");
                        PP_res2.AllowNone = false;
                        PP_res2.UseBasePoint = true;
                        PP_res2.BasePoint = Point_result1.Value;
                        Point_result2 = Editor1.GetPoint(PP_res2);

                        if (Point_result2.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_result3;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_res3;
                        PP_res3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the third point");
                        PP_res3.AllowNone = false;
                        PP_res3.UseBasePoint = true;
                        PP_res3.BasePoint = Point_result1.Value;
                        Point_result3 = Editor1.GetPoint(PP_res3);

                        if (Point_result3.Status != PromptStatus.OK)
                        {
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        Point3d Point1 = Poly1.GetClosestPointTo(Point_result1.Value, Vector3d.ZAxis, false);
                        Point3d Point2 = Poly1.GetClosestPointTo(Point_result2.Value, Vector3d.ZAxis, false);
                        Point3d Point3 = Poly1.GetClosestPointTo(Point_result3.Value, Vector3d.ZAxis, false);



                        double Param1 = Poly1.GetParameterAtPoint(Point1);
                        double Param2 = Poly1.GetParameterAtPoint(Point2);
                        double Param3 = Poly1.GetParameterAtPoint(Point3);

                        double Pc1 = Param1;
                        double Pc2 = Param2;
                        double Pc3 = Param3;

                        if ((Pc1 > Pc2 & Pc2 > Pc3) | (Pc1 < Pc2 & Pc2 < Pc3))
                        {
                            if (Pc1 > Pc2 & Pc2 > Pc3)
                            {
                                double T = Pc1;
                                Pc1 = Pc3;
                                Pc3 = T;
                                T = Param1;
                                Param1 = Param3;
                                Param3 = T;
                                Point3d Pt = Point1;
                                Point1 = Point3;
                                Point3 = Pt;

                            }

                            if (Math.Ceiling(Param1) == Param1)
                            {
                                Pc1 = Param1 + 0.5;
                            }

                            if (Math.Floor(Param3) == Param3)
                            {
                                Pc3 = Param3 - 0.5;
                            }

                            if (Pc3 - Pc1 > 1)
                            {
                                for (int i = Convert.ToInt32(Math.Floor(Pc3)); i > Pc1; --i)
                                {
                                    Poly1.RemoveVertexAt(i);
                                }

                                int Nextv = Convert.ToInt32(Math.Ceiling(Param1));

                                if (Pc1 == Param1)
                                {
                                    Poly1.AddVertexAt(Nextv, new Point2d(Point1.X, Point1.Y), 0, Poly1.GetStartWidthAt(0), Poly1.GetStartWidthAt(0));
                                }
                                if (Pc3 == Param3)
                                {
                                    Poly1.AddVertexAt(Nextv + 1, new Point2d(Point3.X, Point3.Y), 0, Poly1.GetStartWidthAt(0), Poly1.GetStartWidthAt(0));
                                }
                            }
                            else
                            {
                                MessageBox.Show("Your picked points don't have vertexes to be removed");
                            }
                        }

                        if ((Pc1 > Pc2 & Pc2 < Pc3 & Pc1 > Pc3) | (Pc2 > Pc1 & Pc3 < Pc1 & Pc3 < Pc2) | (Pc1 > Pc2 & Pc2 < Pc3 & Pc1 < Pc3) | (Pc2 > Pc3 & Pc2 > Pc1 & Pc3 > Pc1))
                        {

                            if ((Pc1 > Pc2 & Pc2 < Pc3 & Pc1 < Pc3) | (Pc2 > Pc3 & Pc2 > Pc1 & Pc3 > Pc1))
                            {
                                double T = Pc1;
                                Pc1 = Pc3;
                                Pc3 = T;
                                T = Param1;
                                Param1 = Param3;
                                Param3 = T;
                                Point3d Pt = Point1;
                                Point1 = Point3;
                                Point3 = Pt;
                            }

                            if (Math.Ceiling(Param1) == Param1)
                            {
                                Pc1 = Param1 + 0.5;
                            }

                            if (Math.Floor(Param3) == Param3)
                            {
                                Pc3 = Param3 - 0.5;
                            }

                            if (Pc1 + 1 < Poly1.NumberOfVertices)
                            {
                                for (int i = Convert.ToInt32(Poly1.NumberOfVertices - 1); i > Pc1; --i)
                                {
                                    Poly1.RemoveVertexAt(i);
                                }
                            }

                            if (Pc3 >0)
                            {
                                for (int i = 0; i < Pc3; ++i)
                                {
                                    Poly1.RemoveVertexAt(i);
                                }
                            }

                            if (Pc1 == Param1)
                            {
                                Poly1.AddVertexAt(Poly1.NumberOfVertices, new Point2d(Point1.X, Point1.Y), 0, Poly1.GetStartWidthAt(0), Poly1.GetStartWidthAt(0));
                            }
                            if (Pc3 == Param3)
                            {
                                Poly1.AddVertexAt(0, new Point2d(Point3.X, Point3.Y), 0, Poly1.GetStartWidthAt(0), Poly1.GetStartWidthAt(0));
                            }
                            Poly1.Closed = true;

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



        [CommandMethod("CLT")]
        public void Show_CLtool_Form()
        {
            if (isSECURE() == true)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is CLTool_Form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        return;
                    }
                }

                try
                {
                    CLTool_Form forma2 = new CLTool_Form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                return;
            }
        }



    }
}



