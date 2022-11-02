using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;




using Microsoft.Office.Interop.Excel;
using System.Data;

namespace Alignment_mdi
{
    class Functions
    {
        public static bool is_dan_popescu()
        {
            if (Environment.UserName.ToUpper() == "POP70694")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool is_richard_pangburn()
        {
            if (Environment.UserName.ToUpper() == "PAN71158")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_hector_morales()
        {
            if (Environment.UserName.ToUpper() == "MOR72937")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool is_eric_st_germain()
        {
            if (Environment.UserName.ToUpper() == "STG46680")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static bool is_acacia_antley()
        {
            if (Environment.UserName.ToUpper() == "ANT37918")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_alex_dumais()
        {
            if (Environment.UserName.ToUpper() == "DUM64749")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool is_joe_lynskey()
        {
            if (Environment.UserName.ToUpper() == "LYN69372")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool is_monica_forgarty()
        {
            if (Environment.UserName.ToUpper() == "WHI45143")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool is_eli_barboza()
        {
            if (Environment.UserName.ToUpper() == "BAR55261")
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        static public bool IsNumeric(string s)
        {
            double myNum = 0;
            if (double.TryParse(s, out myNum))
            {
                if (s.Contains(",")) return false;
                return true;
            }
            else
            {
                return false;
            }
        }

        static public void Creaza_layer(string Layername, short Culoare, bool Plot)
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
                        if (LayerTable1.Has(Layername) == true)
                        {
                            LayerTable1.UpgradeOpen();
                            LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                            if (new_layer != null)
                            {
                                new_layer.IsPlottable = Plot;

                            }
                        }

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

        static public double GET_Bearing_rad(double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
        }

        static public double GET_deltaX_rad()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return CurentUCS.Xaxis.AngleOnPlane(Planul_curent);
        }

        static public string Get_Quadrant_bearing(double Radian1)
        {
            string Prefix1 = "N ";
            string Suffix1 = " E";
            double Quadrant1 = Math.PI / 2 - Radian1;
            if (Radian1 > Math.PI / 2 & Radian1 <= Math.PI)
            {
                Quadrant1 = Radian1 - Math.PI / 2;
                Suffix1 = " W";
            }
            if (Radian1 > Math.PI & Radian1 <= 3 * Math.PI / 2)
            {
                Quadrant1 = 3 * Math.PI / 2 - Radian1;
                Prefix1 = "S ";
                Suffix1 = " W";
            }
            if (Radian1 > 3 * Math.PI / 2)
            {
                Quadrant1 = Radian1 - 3 * Math.PI / 2;
                Prefix1 = "S ";
                Suffix1 = " E";
            }
            return Prefix1 + Get_DMS(Quadrant1 * 180 / Math.PI, 0) + Suffix1;
        }

        static public string Get_DMS(double Numar, int round_seconds)
        {

            bool Negative = false;
            if (Numar < 0)
            {
                Negative = true;
                Numar = -Numar;
            }
            int Degree1 = Convert.ToInt32(Math.Floor(Numar));



            int Minutes1 = Convert.ToInt32(Math.Floor((Numar - Convert.ToDouble(Degree1)) * 60));

            double rest1 = Convert.ToDouble(Degree1) + Convert.ToDouble(Minutes1) / 60;
            double Seconds1 = Math.Round((Numar - rest1) * 3600, round_seconds);



            if (Seconds1 == 60)
            {
                Minutes1 = Minutes1 + 1;
                Seconds1 = 0;
            }

            if (Minutes1 == 60)
            {
                Degree1 = Degree1 + 1;
                Minutes1 = 0;
            }

            string D = Degree1.ToString();
            if (D.Length == 1) D = "0" + D;

            if (Negative == true) D = "-" + D;

            string M = Minutes1.ToString();
            string S = Get_String_Rounded(Seconds1, round_seconds);

            if (M.Length == 1)
            {
                M = "0" + M;
            }

            if (Seconds1 < 10)
            {
                S = "0" + S;
            }

            char deg_symbol = (char)176;
            char sec_symbol = (char)34;

            return D + deg_symbol + M + "'" + S + sec_symbol;
        }

        static public BlockReference InsertBlock_with_multiple_atributes_with_database(Database Database1, Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord,
            string Nume_fisier, string NumeBlock, Point3d Insertion_point, double Scale_xyz, double Rotation1, string Layer1,
             System.Collections.Specialized.StringCollection Colectie_nume_atribute, System.Collections.Specialized.StringCollection Colectie_valori_atribute)
        {

            BlockReference Block1 = null;


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                if (BlockTable1.Has(NumeBlock) == false)
                {
                    if (System.IO.File.Exists(Nume_fisier) == true)
                    {
                        using (Database Database2 = new Database(false, false))
                        {
                            Database2.ReadDwgFile(Nume_fisier, System.IO.FileShare.Read, true, null);
                            Database1.Insert(NumeBlock, Database2, false);
                        }
                    }


                }

                Trans1.Commit();
            }

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                if (BlockTable1.Has(NumeBlock) == true)
                {


                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTR = (BlockTableRecord)Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForRead);

                    Block1 = new BlockReference(Insertion_point, BTR.ObjectId);
                    Block1.Layer = Layer1;
                    Block1.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz);
                    Block1.Rotation = Rotation1;
                    BTrecord.AppendEntity(Block1);
                    Trans1.AddNewlyCreatedDBObject(Block1, true);
                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;
                    BlockTableRecordEnumerator BTR_enum = BTR.GetEnumerator();
                    while (BTR_enum.MoveNext())
                    {
                        Entity Ent1 = (Entity)Trans1.GetObject(BTR_enum.Current, OpenMode.ForWrite);
                        if (Ent1 is AttributeDefinition)
                        {
                            AttributeDefinition Attdef = (AttributeDefinition)Ent1;
                            AttributeReference Attref = new AttributeReference();
                            Attref.SetAttributeFromBlock(Attdef, Block1.BlockTransform);

                            for (int i = 0; i < Colectie_nume_atribute.Count; ++i)
                            {
                                string Tag1 = Colectie_nume_atribute[i];
                                string Valoare = Colectie_valori_atribute[i];
                                if (Attref.Tag.ToLower() == Tag1.ToLower())
                                {
                                    Attref.TextString = Valoare;
                                    i = Colectie_nume_atribute.Count;
                                }
                            }
                            if (Attref != null)
                            {
                                attColl.AppendAttribute(Attref);
                                Trans1.AddNewlyCreatedDBObject(Attref, true);
                            }
                        }

                    }

                }

                Trans1.Commit();
            }

            return Block1;
        }
        static public void Stretch_block(BlockReference BR, String Prop_name, double Prop_value)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    if (prop.PropertyName == Prop_name && prop.UnitsType == DynamicBlockReferencePropertyUnitsType.Distance)
                    {
                        prop.Value = Prop_value;
                        return;
                    }
                }
            }
        }

        static public string Get_String_Rounded(double Numar, int Nr_dec)
        {

            String String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";

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

            String1 = Math.Round(Numar, Nr_dec, MidpointRounding.AwayFromZero).ToString();

            String2 = String1;

            if (String1.Contains(".") == false)
            {
                String2 = String1 + String_punct + Zero;
                goto end;
            }

            if (String1.Length - String1.IndexOf(".") - 1 - Nr_dec != 0)
            {
                for (int i = 1; i <= String1.IndexOf(".") + 1 + Nr_dec - String1.Length; i = i + 1)
                {
                    zero1 = zero1 + "0";
                }

                String2 = String1 + zero1;
            }

            end:
            return String_minus + String2;

        }

        public static System.Data.DataTable Creaza_lgen_alias_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("Layer name", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Label Layer", typeof(string));
            dt.Columns.Add("Boundary_Layer (Yes/No)", typeof(string));
            dt.Columns.Add("Align_To_Feature (Yes/No)", typeof(string));
            dt.Columns.Add("Primary_Label_Type", typeof(string));
            dt.Columns.Add("Secondary_Label_Type", typeof(string));
            dt.Columns.Add("Tertiary_Label_Type", typeof(string));
            dt.Columns.Add("Mtext_Style_Name", typeof(string));
            dt.Columns.Add("Mtext_Style_Font", typeof(string));
            dt.Columns.Add("Mtext_Style_Width_Factor", typeof(string));
            dt.Columns.Add("Mtext_Style_Oblique_Angle", typeof(string));
            dt.Columns.Add("Mtext_Style_Height_1:1", typeof(string));
            dt.Columns.Add("Mtext_Underline (Yes/No)", typeof(string));
            dt.Columns.Add("Mleader Style Name", typeof(string));
            dt.Columns.Add("Mleader Arrow size", typeof(string));
            dt.Columns.Add("Mleader Gap", typeof(string));
            dt.Columns.Add("Mleader Dog Length", typeof(string));
            dt.Columns.Add("Mleader Text height at 1:1", typeof(string));
            dt.Columns.Add("Use_Object_Data (Yes/No)", typeof(string));
            dt.Columns.Add("Break lines (Yes/No)", typeof(string));
            dt.Columns.Add("Force_Caps (Yes/No)", typeof(string));
            dt.Columns.Add("Contour_layer (Yes/No)", typeof(string));
            dt.Columns.Add("Contour_Label_precision", typeof(string));
            dt.Columns.Add("Prefix1", typeof(string));
            dt.Columns.Add("Object Data Field1", typeof(string));
            dt.Columns.Add("Suffix1", typeof(string));
            dt.Columns.Add("Prefix2", typeof(string));
            dt.Columns.Add("Object Data Field2", typeof(string));
            dt.Columns.Add("Suffix2", typeof(string));
            dt.Columns.Add("Prefix3", typeof(string));
            dt.Columns.Add("Object Data Field3", typeof(string));
            dt.Columns.Add("Suffix3", typeof(string));
            dt.Columns.Add("Prefix4", typeof(string));
            dt.Columns.Add("Object Data Field4", typeof(string));
            dt.Columns.Add("Suffix4", typeof(string));
            dt.Columns.Add("Prefix5", typeof(string));
            dt.Columns.Add("Object Data Field5", typeof(string));
            dt.Columns.Add("Suffix5", typeof(string));
            dt.Columns.Add("Prefix6", typeof(string));
            dt.Columns.Add("Object Data Field6", typeof(string));
            dt.Columns.Add("Suffix6", typeof(string));
            dt.Columns.Add("Prefix7", typeof(string));
            dt.Columns.Add("Object Data Field7", typeof(string));
            dt.Columns.Add("Suffix7", typeof(string));
            dt.Columns.Add("Prefix8", typeof(string));
            dt.Columns.Add("Object Data Field8", typeof(string));
            dt.Columns.Add("Suffix8", typeof(string));
            dt.Columns.Add("Prefix9", typeof(string));
            dt.Columns.Add("Object Data Field9", typeof(string));
            dt.Columns.Add("Suffix9", typeof(string));
            dt.Columns.Add("Prefix10", typeof(string));
            dt.Columns.Add("Object Data Field10", typeof(string));
            dt.Columns.Add("Suffix10", typeof(string));
            dt.Columns.Add("Block Name", typeof(string));
            dt.Columns.Add("Block Attribute using concatenated description", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_1", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_2", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_3", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_4", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_5", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_6", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_7", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_8", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_9", typeof(string));
            dt.Columns.Add("Block_Attribute_Source_OD_Field_10", typeof(string));
            dt.Columns.Add("DimStyle_name", typeof(string));
            dt.Columns.Add("DimStyle_ArrowSize", typeof(string));
            dt.Columns.Add("DimStyle_Suffix", typeof(string));
            dt.Columns.Add("DimStyle_Decimals_no", typeof(string));
            dt.Columns.Add("DimStyle_Round_to_closest", typeof(string));
            dt.Columns.Add("DimStyle_force_dimline", typeof(string));
            dt.Columns.Add("Background Mask (Yes/No)", typeof(string));
            dt.Columns.Add("Text Frame Mleaders Only (Yes/No)", typeof(string));
            dt.Columns.Add("Precision (Decimal Places)", typeof(string));
            return dt;
        }

        public static System.Data.DataTable Build_Lgen_Data_table_layer_alias_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {


            System.Data.DataTable Data_table_alias = Creaza_lgen_alias_datatable_structure();
            int NrR = 0;
            int NrC = Data_table_alias.Columns.Count;

            string col1 = "A";



            for (int i = Start_row; i < 30000; ++i)
            {
                if (i == Start_row)
                {
                    if (W1.Range[col1 + i.ToString()].Value2 == null)
                    {
                        return Data_table_alias;
                    }
                }

                if (W1.Range[col1 + i.ToString()].Value2 == null)
                {
                    NrR = i - Start_row;
                    i = 31000;
                }
                else
                {
                    Data_table_alias.Rows.Add();
                }
            }


            Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];





            object[,] values = new object[NrR - 1, NrC - 1];

            values = range1.Value2;

            for (int i = 0; i < Data_table_alias.Rows.Count; ++i)
            {
                for (int j = 0; j < Data_table_alias.Columns.Count; ++j)
                {
                    object Valoare = values[i + 1, j + 1];
                    if (Valoare == null) Valoare = DBNull.Value;

                    Data_table_alias.Rows[i][j] = Valoare;
                }
            }




            return Data_table_alias;


        }

        static public string Get_String_Rounded_with_thousand_sep(double Numar, int Nr_dec)
        {

            string String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";
            string Comma = ",";

            string String_punct = "";

            if (Nr_dec > 0)
            {
                String_punct = ".";
                for (int i = 1; i <= Nr_dec; i = i + 1)
                {
                    Zero = Zero + "0";
                }
            }

            string String_minus = "";

            if (Numar < 0)
            {
                String_minus = "-";
                Numar = -Numar;
            }


            double Numar_double = Math.Round(Numar, Nr_dec, MidpointRounding.AwayFromZero);
            string Numar_int = Math.Floor(Numar_double).ToString();

            String1 = Numar_double.ToString();


            if (Numar_double >= 1000)
            {

                string Rest1 = Numar_int;
                String String3 = "";

                int No_of_1000 = Convert.ToInt32(Math.Floor(Convert.ToDouble(Numar_int.Length / 3)));
                int Multiple_rest = Numar_int.Length - No_of_1000 * 3;
                if (Multiple_rest > 0)
                {
                    String3 = Numar_int.Substring(0, Multiple_rest) + Comma;
                    Rest1 = Numar_int.Substring(Multiple_rest, Numar_int.Length - Multiple_rest);
                }

                for (int i = 0; i < No_of_1000; ++i)
                {
                    String3 = String3 + Rest1.Substring(i * 3, 3) + Comma;
                }

                String3 = String3.Substring(0, String3.Length - 1);
                double Multiplier = 1;

                if (Nr_dec > 0)
                {
                    for (int i = 0; i < Nr_dec; ++i)
                    {
                        Multiplier = Multiplier * 10;
                    }

                }


                double Diferenta = Math.Round(Numar_double - Math.Floor(Numar_double), Nr_dec) * Multiplier;
                string String4 = Diferenta.ToString();

                if (Nr_dec > 0)
                {
                    if (String4.Length < Nr_dec)
                    {
                        for (int i = 0; i < Nr_dec - String4.Length; ++i)
                        {
                            String4 = "0" + String4;
                        }
                    }
                    return String_minus + String3 + String_punct + String4;
                }

                if (Nr_dec == 0) return String_minus + String3;


            }


            String2 = String1;

            if (String1.Contains(".") == false)
            {
                String2 = String1 + String_punct + Zero;
                goto end;
            }

            if (String1.Length - String1.IndexOf(".") - 1 - Nr_dec != 0)
            {
                for (int i = 1; i <= String1.IndexOf(".") + 1 + Nr_dec - String1.Length; i = i + 1)
                {
                    zero1 = zero1 + "0";
                }

                String2 = String1 + zero1;
            }

            end:
            return String_minus + String2;

        }

        public static Matrix3d ModelToPaper(Viewport vp)
        {
            Vector3d vd = vp.ViewDirection;
            Point3d vc = new Point3d(vp.ViewCenter.X, vp.ViewCenter.Y, 0);
            Point3d vt = vp.ViewTarget;
            Point3d cp = vp.CenterPoint;
            double ta = -vp.TwistAngle;
            double vh = vp.ViewHeight;
            double height = vp.Height;
            double width = vp.Width;
            double scale = vh / height;
            double lensLength = vp.LensLength;
            Vector3d zaxis = vd.GetNormal();
            Vector3d xaxis = Vector3d.ZAxis.CrossProduct(vd);
            Vector3d yaxis;

            if (!xaxis.IsZeroLength())
            {
                xaxis = xaxis.GetNormal();
                yaxis = zaxis.CrossProduct(xaxis);
            }
            else if (zaxis.Z < 0)
            {
                xaxis = Vector3d.XAxis * -1;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis * -1;
            }
            else
            {
                xaxis = Vector3d.XAxis;
                yaxis = Vector3d.YAxis;
                zaxis = Vector3d.ZAxis;
            }
            Matrix3d pcsToDCS = Matrix3d.Displacement(Point3d.Origin - cp);
            pcsToDCS = pcsToDCS * Matrix3d.Scaling(scale, cp);
            Matrix3d dcsToWcs = Matrix3d.Displacement(vc - Point3d.Origin);
            Matrix3d mxCoords = Matrix3d.AlignCoordinateSystem(
                  Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis,
                 Vector3d.ZAxis, Point3d.Origin,
                xaxis, yaxis, zaxis);
            dcsToWcs = mxCoords * dcsToWcs;
            dcsToWcs = Matrix3d.Displacement(vt - Point3d.Origin) * dcsToWcs;
            dcsToWcs = Matrix3d.Rotation(ta, zaxis, vt) * dcsToWcs;

            Matrix3d perspectiveMx = Matrix3d.Identity;
            if (vp.PerspectiveOn)
            {
                double vSize = vh;
                double aspectRatio = width / height;
                double adjustFactor = 1.0 / 42.0;
                double adjstLenLgth = vSize * lensLength *
                   Math.Sqrt(1.0 + aspectRatio * aspectRatio) * adjustFactor;
                double iDist = vd.Length;
                double lensDist = iDist - adjstLenLgth;
                double[] dataAry = new double[]
                {
                     1,0,0,0,0,1,0,0,0,0,
                       (adjstLenLgth-lensDist)/adjstLenLgth,
                       lensDist*(iDist-adjstLenLgth)/adjstLenLgth,
                     0,0,-1.0/adjstLenLgth,iDist/adjstLenLgth
               };

                perspectiveMx = new Matrix3d(dataAry);
            }

            Matrix3d finalMx =
             pcsToDCS.Inverse() * perspectiveMx * dcsToWcs.Inverse();

            return finalMx;
        }
        public static Matrix3d PaperToModel(Viewport vp)
        {
            Matrix3d mx = ModelToPaper(vp);
            return mx.Inverse();
        }


    }

}
