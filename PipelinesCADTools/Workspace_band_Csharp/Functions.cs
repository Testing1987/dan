using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace Workspace_Band
{
    public class Functions
    {
        static public Worksheet Get_active_worksheet_from_Excel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;
                Workbook1 = Excel1.ActiveWorkbook;
                return Workbook1.ActiveSheet;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }


        }


        static public bool IsNumeric(string s)
        {
            double myNum = 0;
            if (Double.TryParse(s, out myNum))
            {
                if (s.Contains(",")) return false;
                return true;
            }
            else
            {
                return false;
            }
        }


        static public double GET_Bearing_rad(Double x1, Double y1, Double x2, Double y2)
        {
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
        }

        public void Incarca_existing_layers_to_combobox(System.Windows.Forms.ComboBox Combo_layer)
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        Combo_layer.Items.Clear();

                        string[] Array1 = null;

                        int idx1 = 1;
                        foreach (ObjectId Layer_id in layer_table)
                        {
                            LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            string Nume_layer = Layer1.Name;
                            if (Nume_layer.Contains("|") == false & Nume_layer.Contains("$") == false)
                            {
                                Array.Resize(ref Array1, idx1);
                                Array1[idx1 - 1] = Nume_layer;
                                idx1 = idx1 + 1;
                            }

                        }
                        if (Array1 != null)
                        {
                            System.Array.Sort(Array1);
                            for (int i = 0; i < Array1.Length; ++i)
                            {
                                Combo_layer.Items.Add(Array1[i]);
                            }


                        }

                        Trans1.Dispose();
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }

        }


        static public String Get_String_Rounded(Double Numar, int Nr_dec)
        {
            String String1;
            String Zero = "";
            String zero1 = "";


            String String_punct = "";

            if (Nr_dec > 0)
            {
                String_punct = ".";
                for (int i = 1; i <= Nr_dec; i = i + 1)
                {
                    Zero = Zero + "0";
                }

            }

            String String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }

            String1 = Convert.ToString(Math.Round(Numar, Nr_dec, MidpointRounding.AwayFromZero));
            String String2 = String1;

            if (String1.Contains(".") == false)
            {
                String2 = String1 + String_punct + Zero;
                goto calcs;
            }

            if (String1.Length - String1.IndexOf(".") + 1 - Nr_dec != 0)
            {
                for (int i = 1; i <= String1.IndexOf(".") + 1 + Nr_dec - String1.Length; i = i + 1)
                {
                    zero1 = zero1 + "0";
                }

                String2 = String1 + zero1;
            }
        calcs:

            return (String_minus + String2);
        }

        static public String Get_chainage_feet_from_double(Double Numar, int Nr_dec)
        {
            String String_minus = "";
            String String2 = "";
            String String3 = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }

            String2 = Get_String_Rounded(Numar, Nr_dec);

            int Punct;
            if (String2.Contains(".") == false)
            {
                Punct = 0;
            }
            else
            {
                Punct = 1;
            }

            if (String2.Length - Nr_dec - Punct >= 4)
            {
                String3 = String2.Substring(0, String2.Length - 2 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
            }
            else
            {
                if (String2.Length - Nr_dec - Punct == 1) String3 = "0+0" + String2;
                if (String2.Length - Nr_dec - Punct == 2) String3 = "0+" + String2;
                if (String2.Length - Nr_dec - Punct == 3) String3 = String2.Substring(0, 1) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));

            }

            return (String_minus + String3);


        }



        static public MLeader Mleader_Create_without_UCS_transform(Point3d Point1, string Continut, double text_height, double arrow_size, double Gap1, double Delta_X, double Delta_Y)
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        MText Mtext1 = new MText();
                        Mtext1.TextHeight = text_height;
                        Mtext1.Contents = Continut;
                        Mtext1.ColorIndex = 0;
                        MLeader Mleader1 = new MLeader();
                        int Nr1 = Mleader1.AddLeader();
                        int Nr2 = Mleader1.AddLeaderLine(Nr1);
                        Mleader1.AddFirstVertex(Nr2, new Point3d(Point1.X, Point1.Y, Point1.Z));
                        Mleader1.AddLastVertex(Nr2, new Point3d(Point1.X + Delta_X, Point1.Y + Delta_Y, Point1.Z));
                        Mleader1.LeaderLineType = LeaderType.StraightLeader;
                        Mleader1.ContentType = ContentType.MTextContent;
                        Mleader1.MText = Mtext1;
                        Mleader1.TextHeight = text_height;
                        Mleader1.LandingGap = Gap1;
                        Mleader1.ArrowSize = arrow_size;
                        Mleader1.DoglegLength = Gap1;
                        Mleader1.Annotative = AnnotativeStates.False;
                        BTrecord.AppendEntity(Mleader1);
                        Trans1.AddNewlyCreatedDBObject(Mleader1, true);
                        Trans1.Commit();
                        return Mleader1;
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
        }


        static public Polyline Clean_poly_of_duplicate_points(Polyline Poly1, int precision)
        {
            try
            {
                Polyline Poly2 = new Polyline();
                Poly2.Elevation = Poly1.Elevation;

                double X = 0;
                double Y = 0;

                int j = 1;

                for (int i = 0; i < Poly1.NumberOfVertices; i = i + 1)
                {
                    if (i == 0)
                    {
                        Poly2.AddVertexAt(0, new Point2d(Poly1.GetPoint2dAt(0).X, Poly1.GetPoint2dAt(0).Y), Poly1.GetBulgeAt(0), Poly1.GetStartWidthAt(0), Poly1.GetEndWidthAt(0));
                        X = Math.Round(Poly1.GetPoint2dAt(0).X, precision);
                        Y = Math.Round(Poly1.GetPoint2dAt(0).Y, precision);
                    }
                    else
                    {
                        Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(X, Y, 0), new Point3d(Math.Round(Poly1.GetPoint2dAt(i).X, precision), Math.Round(Poly1.GetPoint2dAt(i).X, precision), 0));

                        if (Line1.Length >= precision/100)
                        {
                            Poly2.AddVertexAt(j, new Point2d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y), Poly1.GetBulgeAt(i), Poly1.GetStartWidthAt(i), Poly1.GetEndWidthAt(i));
                            X = Math.Round(Poly1.GetPoint2dAt(i).X, precision);
                            Y = Math.Round(Poly1.GetPoint2dAt(i).Y, precision);
                            j = j + 1;
                        }
                    }
                }

                return Poly2;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
        }
       
        static public Polyline Clean_poly_of_deflection_points(Polyline Poly1, double min_Deflection_rad)
        {
            try
            {
                Polyline Poly2 = new Polyline();
                Poly2.Elevation = Poly1.Elevation;

                double X1 = 0;
                double Y1 = 0;
                double X2 = 0;
                double Y2 = 0;
                double X3 = 0;
                double Y3 = 0;

                int j = 1;

                for (int i = 0; i < Poly1.NumberOfVertices; i = i + 1)
                {
                    if (i == 0)
                    {
                        Poly2.AddVertexAt(0, new Point2d(Poly1.GetPoint2dAt(0).X, Poly1.GetPoint2dAt(0).Y), Poly1.GetBulgeAt(0), Poly1.GetStartWidthAt(0), Poly1.GetEndWidthAt(0));
                        X1 = Poly1.GetPoint2dAt(0).X;
                        Y1 = Poly1.GetPoint2dAt(0).Y;
                    }

                    else if (i < Poly1.NumberOfVertices-1)
                    {
                        X2 = Poly1.GetPoint2dAt(i).X;
                        Y2 = Poly1.GetPoint2dAt(i).Y;

                        X3 = Poly1.GetPoint2dAt(i+1).X;
                        Y3 = Poly1.GetPoint2dAt(i+1).Y;

                        Vector3d vector1 = new Point3d(X1, Y1, Poly1.Elevation).GetVectorTo(new Point3d(X2, Y2, Poly1.Elevation));
                        Vector3d vector2 = new Point3d(X2, Y2, Poly1.Elevation).GetVectorTo(new Point3d(X3, Y3, Poly1.Elevation));
                        
                        double Defl = Math.Round(vector2.GetAngleTo(vector1), 2);
                        if (Defl >= min_Deflection_rad)
                        {
                            Poly2.AddVertexAt(j, new Point2d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y), Poly1.GetBulgeAt(i), Poly1.GetStartWidthAt(i), Poly1.GetEndWidthAt(i));
                            X1 = Poly1.GetPoint2dAt(i).X;
                            Y1 = Poly1.GetPoint2dAt(i).Y;
                            j = j + 1;
                        }
                    }
                    else if (i == Poly1.NumberOfVertices - 1)
                    {
                        Poly2.AddVertexAt(j, new Point2d(Poly1.GetPoint2dAt(i).X, Poly1.GetPoint2dAt(i).Y), Poly1.GetBulgeAt(i), Poly1.GetStartWidthAt(i), Poly1.GetEndWidthAt(i));
                    }


                }

                return Poly2;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }
        }

    }
}
