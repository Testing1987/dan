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
        static public Point3dCollection Intersect_on_both_operands(Curve Curba1, Curve Curba2)
        {
            Point3dCollection Col_int = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands = new Point3dCollection();
            Point3dCollection Col_int_on_both_operands_DUPLICATE = new Point3dCollection();

            Curba1.IntersectWith(Curba2, Intersect.OnBothOperands, Col_int, IntPtr.Zero, IntPtr.Zero);

            if (Col_int.Count == 1) return Col_int;
            if (Col_int.Count == 0) return Col_int;

            if (Col_int.Count > 1)
            {
                if (Curba1 is Polyline & Curba2 is Polyline)
                {
                    for (int i = 0; i < Col_int.Count; ++i)
                    {
                        Point3d Pt1 = new Point3d();
                        Pt1 = Col_int[i];
                        try
                        {
                            double param_on_1 = Curba1.GetParameterAtPoint(Pt1);
                            double param_on_2 = Curba2.GetParameterAtPoint(Pt1);


                            if (Col_int_on_both_operands_DUPLICATE.Contains(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4))) == false)
                            {
                                Col_int_on_both_operands.Add(Pt1);
                                Col_int_on_both_operands_DUPLICATE.Add(new Point3d(Math.Round(Pt1.X, 4), Math.Round(Pt1.Y, 4), Math.Round(Pt1.Z, 4)));
                            }
                        }
                        catch (System.Exception ex)
                        {
                        }
                    }
                    return Col_int_on_both_operands;
                }
                else
                {
                    return Col_int;
                }
            }
            else
            {
                return Col_int;
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


        static public Worksheet Get_NEW_worksheet_from_Excel()
        {
            Microsoft.Office.Interop.Excel.Application Excel1;
            Microsoft.Office.Interop.Excel.Workbook Workbook1;
            try
            {
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Exception ex)
            {
                Excel1 = new Microsoft.Office.Interop.Excel.Application();
            }

            try
            {
                Excel1.Visible = true;
                Excel1.Workbooks.Add();
                Workbook1 = Excel1.ActiveWorkbook;
                return Workbook1.ActiveSheet;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }


        }
        static public int Get_no_of_workbooks_from_Excel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return 0;
                return Excel1.Workbooks.Count;
            }
            catch (System.Exception ex)
            {
                return 0;
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
        static public string Get_chainage_from_double(double Numar, string units, int Nr_dec)
        {

            String String2, String3;
            String3 = "";
            String String_minus = "";

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
                if (units == "f") String3 = String2.Substring(0, String2.Length - 2 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                if (units == "m") String3 = String2.Substring(0, String2.Length - 3 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (3 + Nr_dec + Punct));
            }
            else
            {
                if (units == "f")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = String2.Substring(0, 1) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
                }
                if (units == "m")
                {
                    if (String2.Length - Nr_dec - Punct == 1) String3 = "0+00" + String2;
                    if (String2.Length - Nr_dec - Punct == 2) String3 = "0+0" + String2;
                    if (String2.Length - Nr_dec - Punct == 3) String3 = "0+" + String2;
                }
            }


            return String_minus + String3;

        }

        static public double Get_deflection_angle_rad(double x1, double y1, double x2, double y2, double x3, double y3)
        {
            double a1 = x2 - x1;
            double b1 = y2 - y1;
            double a2 = x3 - x2;
            double b2 = y3 - y2;
            double Defl_DD = Math.Acos((a1 * a2 + b1 * b2) / (Math.Pow(a1 * a1 + b1 * b1, 0.5) * Math.Pow(a2 * a2 + b2 * b2, 0.5)));
            //return Defl_DD;

            Vector3d vector1 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));
            Vector3d vector2 = new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x3, y3, 0));
            return (vector2.GetAngleTo(vector1));


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


        public static System.Data.DataTable Creaza_weldmap_pipelist_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Pipe ID", typeof(string));
            dt.Columns.Add("Heat", typeof(string));
            dt.Columns.Add("Length", typeof(string));
            dt.Columns.Add("Wall Thickness", typeof(string));
            dt.Columns.Add("Diameter", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Coating", typeof(string));
            dt.Columns.Add("Manufacture", typeof(string));
            dt.Columns.Add("DoubleJointNo", typeof(string));
            return dt;
        }
        public static System.Data.DataTable Creaza_weldmap_pipe_tally_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("MMID", typeof(string));
            dt.Columns.Add("Pipe", typeof(string));
            dt.Columns.Add("Heat", typeof(string));
            dt.Columns.Add("OriginalLength", typeof(string));
            dt.Columns.Add("NewLength", typeof(string));
            dt.Columns.Add("WallThickness", typeof(string));
            dt.Columns.Add("Diameter", typeof(string));
            dt.Columns.Add("Grade", typeof(string));
            dt.Columns.Add("Coating", typeof(string));
            dt.Columns.Add("Manufacture", typeof(string));
            dt.Columns.Add("DoubleJointNo", typeof(string));
            return dt;
        }

        static public System.Data.DataTable Populate_data_table_from_excel(System.Data.DataTable dt1, Worksheet W1, int start_row,
            string checkColumn1, string checkColumn2, string checkColumn3, string checkColumn4, string checkColumn5, string checkColumn6, string checkColumn7, string checkColumn8, string checkColumn9, string checkColumn10, string checkColumn11,
            bool show_message)
        {
            if (W1 == null) return dt1;


            if (checkColumn1 != "")
            {
                Range range1 = W1.Range[checkColumn1 + start_row.ToString() + ":" + checkColumn1 + "30000"];
                object[,] values2 = new object[300000, 1];
                values2 = range1.Value2;


                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (Valoare2 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values2.Length + 1;
                    }
                }
            }

            if (checkColumn2 != "")
            {
                Range range2 = W1.Range[checkColumn2 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn2 + "30000"];
                object[,] values3 = new object[300000, 1];
                values3 = range2.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values3.Length + 1;
                    }
                }
            }

            if (checkColumn3 != "")
            {
                Range range3 = W1.Range[checkColumn3 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn3 + "30000"];
                object[,] values3 = new object[300000, 1];
                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values3.Length + 1;
                    }
                }
            }

            if (checkColumn4 != "")
            {
                Range range4 = W1.Range[checkColumn4 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn4 + "30000"];
                object[,] values4 = new object[300000, 1];
                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (Valoare4 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values4.Length + 1;
                    }
                }
            }

            if (checkColumn5 != "")
            {
                Range range5 = W1.Range[checkColumn5 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn5 + "30000"];
                object[,] values5 = new object[300000, 1];
                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (Valoare5 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values5.Length + 1;
                    }
                }
            }

            if (checkColumn6 != "")
            {
                Range range6 = W1.Range[checkColumn6 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn6 + "30000"];
                object[,] values6 = new object[300000, 1];
                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (Valoare6 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values6.Length + 1;
                    }
                }
            }

            if (checkColumn7 != "")
            {
                Range range7 = W1.Range[checkColumn7 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn7 + "30000"];
                object[,] values7 = new object[300000, 1];
                values7 = range7.Value2;

                for (int i = 1; i <= values7.Length; ++i)
                {
                    object Valoare7 = values7[i, 1];
                    if (Valoare7 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values7.Length + 1;
                    }
                }
            }

            if (checkColumn8 != "")
            {
                Range range8 = W1.Range[checkColumn8 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn8 + "30000"];
                object[,] values8 = new object[300000, 1];
                values8 = range8.Value2;

                for (int i = 1; i <= values8.Length; ++i)
                {
                    object Valoare8 = values8[i, 1];
                    if (Valoare8 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values8.Length + 1;
                    }
                }
            }

            if (checkColumn9 != "")
            {
                Range range9 = W1.Range[checkColumn9 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn9 + "30000"];
                object[,] values9 = new object[300000, 1];
                values9 = range9.Value2;

                for (int i = 1; i <= values9.Length; ++i)
                {
                    object Valoare9 = values9[i, 1];
                    if (Valoare9 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values9.Length + 1;
                    }
                }
            }

            if (checkColumn10 != "")
            {
                Range range10 = W1.Range[checkColumn10 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn10 + "30000"];
                object[,] values10 = new object[300000, 1];
                values10 = range10.Value2;

                for (int i = 1; i <= values10.Length; ++i)
                {
                    object Valoare10 = values10[i, 1];
                    if (Valoare10 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values10.Length + 1;
                    }
                }
            }

            if (checkColumn11 != "")
            {
                Range range11 = W1.Range[checkColumn11 + Convert.ToString(start_row + dt1.Rows.Count) + ":" + checkColumn11 + "30000"];
                object[,] values11 = new object[300000, 1];
                values11 = range11.Value2;

                for (int i = 1; i <= values11.Length; ++i)
                {
                    object Valoare11 = values11[i, 1];
                    if (Valoare11 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values11.Length + 1;
                    }
                }
            }

            int NrC = dt1.Columns.Count;
            int NrR = dt1.Rows.Count;

            if (dt1.Rows.Count == 0)
            {
                if (show_message == true)
                {
                    MessageBox.Show("no data found in the file");
                }

                return dt1;
            }

            if (dt1.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[start_row, 1], W1.Cells[NrR + start_row - 1, NrC]];
                object[,] values = new object[NrR - 1, NrC - 1];
                values = range1.Value2;
                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    for (int j = 0; j < dt1.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dt1.Rows[i][j] = Valoare;
                    }
                }
            }

            return dt1;
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

        public static void Populate_object_data_table_from_objectid(Autodesk.Gis.Map.ObjectData.Tables Tables1, ObjectId id1, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {

                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                    {
                        if (Records1.Count > 0)
                        {
                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                            {
                                Autodesk.Gis.Map.Utilities.MapValue Valoare1;
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Valoare1 = Record1[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        Valoare1.Assign(List_value[i].ToString());
                                    }

                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        Valoare1.Assign(Convert.ToDouble(List_value[i]));
                                    }
                                    Records1.UpdateRecord(Record1);
                                }
                            }
                        }
                        else
                        {
                            using (Autodesk.Gis.Map.ObjectData.Record rec = Autodesk.Gis.Map.ObjectData.Record.Create())
                            {
                                Tabla1.InitRecord(rec);
                                for (int i = 0; i < List_value.Count; ++i)
                                {
                                    Autodesk.Gis.Map.Utilities.MapValue Val = rec[i];
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character)
                                    {
                                        string Valoare = List_value[i].ToString();
                                        Val.Assign(Valoare);
                                    }
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        double Valoare = Convert.ToDouble(List_value[i]);
                                        Val.Assign(Valoare);
                                    }
                                }
                                Tabla1.AddRecord(rec, id1);
                            }
                        }
                    }
                }
            }
        }

        static public void zoom_to_Point(Point3d pt, double factor1)
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
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - factor1, pt.Y - factor1, 0);
                            Point3d maxx = new Point3d(pt.X + factor1, pt.Y + factor1, 0);

                            using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                            {

                                int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                //from here 2015 dlls:
                                Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                // to here 2015 dlls

                                //from here 2013 dlls:

                                //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                // to here 2013 dlls

                                if (view != null)
                                {
                                    using (view)
                                    {

                                        view.ZoomExtents(minx, maxx);

                                        view.Zoom(0.95);//<--optional 
                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }
                                }
                                Trans1.Commit();
                            }


                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


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

        static public Worksheet Get_opened_worksheet_from_Excel_by_name(string filename, string SheetName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return null;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    if (Workbook1.Name == filename)
                    {
                        for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                        {
                            if (Workbook1.Worksheets[i].name == SheetName)
                            {
                                return Workbook1.Worksheets[i];
                            }
                        }
                    }
                }
                return null;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
            }



        }

        static public void Load_opened_worksheets_to_combobox(ComboBox combo1)
        {
            combo1.Items.Clear();
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    string wn = Workbook1.Name;
                    for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    {
                        combo1.Items.Add("[" + Workbook1.Worksheets[i].name + "] - " + wn);
                    }
                }
                if (combo1.Items.Count > 0) combo1.SelectedIndex = 0;

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static void Transfer_datatable_to_existing_excel_spreadsheet_by_name(System.Data.DataTable dt1, string filename, string sheetname, bool delete_columns, int startrow, int endrow)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application Excel1;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1;
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    if (Excel1 == null) return;
                    for (int k = 1; k <= Excel1.Workbooks.Count; ++k)
                    {
                        Workbook1 = Excel1.Workbooks[k];
                        string wn = Workbook1.Name;
                        if (wn.ToUpper() == filename.ToUpper())
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                            {
                                if (W1.Name.ToUpper() == sheetname.ToUpper())
                                {
                                    int maxRows = dt1.Rows.Count;
                                    int maxCols = dt1.Columns.Count;
                                    if (delete_columns == true)
                                    {
                                        W1.Columns["A:XX"].Delete();
                                        W1.Cells.NumberFormat = "General";
                                    }
                                    else
                                    {
                                        W1.Rows[startrow.ToString() + ":" + endrow.ToString()].ClearContents();
                                    }


                                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                                    object[,] values1 = new object[maxRows, maxCols];
                                    for (int i = 0; i < maxRows; ++i)
                                    {
                                        for (int j = 0; j < maxCols; ++j)
                                        {
                                            if (dt1.Rows[i][j] != DBNull.Value)
                                            {
                                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                                            }
                                        }
                                    }
                                    for (int i = 0; i < dt1.Columns.Count; ++i)
                                    {
                                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName.ToUpper();
                                    }
                                    range1.Value2 = values1;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        public static void erase_rows_from_excel(string filename, string sheetname, List<int> lista1, int startrow)
        {
            if (lista1 != null)
            {
                if (lista1.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Application Excel1;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1;
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    if (Excel1 == null) return;
                    lista1.Sort();
                    for (int k = 1; k <= Excel1.Workbooks.Count; ++k)
                    {
                        Workbook1 = Excel1.Workbooks[k];
                        string wn = Workbook1.Name;
                        if (wn.ToUpper() == filename.ToUpper())
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                            {
                                if (W1.Name.ToUpper() == sheetname.ToUpper())
                                {

                                    for (int i = lista1.Count - 1; i >= 0; --i)
                                    {
                                        W1.Rows[(startrow + lista1[i]).ToString()].Delete();
                                    }
                                    return;
                                }
                            }
                        }
                    }
                }
            }
        }

        public static Autodesk.Gis.Map.ObjectData.Table Get_object_data_table(string Nume_table, string Description_table, List<string> List_Names, List<string> List_descriptions, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                return Tables1[Nume_table];
            }

            if (Tables1.IsTableDefined(Nume_table) == false)
            {
                using (Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_definitions = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.MapUtility.NewODFieldDefinitions())
                {
                    for (int i = 0; i < List_Names.Count; ++i)
                    {
                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_definitions.Add(List_Names[i], List_descriptions[i], List_types[i], i);
                    }

                    Tables1.Add(Nume_table, Field_definitions, Description_table, true);
                }
            }
            return Tables1[Nume_table];
        }



        public static void Transfer_datatable_to_new_excel_spreadsheet_named(System.Data.DataTable dt1, string nume1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "@";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                    if (nume1 != "")
                    {
                        W1.Name = nume1;
                    }

                }
            }
        }


        public static System.Data.DataTable Sort_data_table(System.Data.DataTable Datatable1, string Column1)
        {
            System.Data.DataTable Data_table_temp = new System.Data.DataTable();
            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    if (Datatable1.Columns.Contains(Column1) == true)
                    {
                        System.Data.DataView DataView1 = new System.Data.DataView(Datatable1);
                        DataView1.Sort = Column1 + " ASC";
                        Data_table_temp = Datatable1.Clone();
                        Data_table_temp.Rows.Clear();
                        for (int i = 0; i < DataView1.Count; ++i)
                        {
                            System.Data.DataRow Data_row1 = DataView1[i].Row;
                            Data_table_temp.Rows.Add();
                            for (int j = 0; j < Datatable1.Columns.Count; ++j)
                            {
                                Data_table_temp.Rows[Data_table_temp.Rows.Count - 1][j] = Data_row1[j];
                            }
                        }
                    }
                }
            }
            return Data_table_temp;

        }

        public static System.Data.DataTable Creaza_weldmap_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("PNT", typeof(string));
            dt.Columns.Add("NORTHING", typeof(string));
            dt.Columns.Add("EASTING", typeof(string));
            dt.Columns.Add("ELEVATION", typeof(string));
            dt.Columns.Add("FEATURE_CODE", typeof(string));
            dt.Columns.Add("DESCRIPTION", typeof(string));
            dt.Columns.Add("STATION_LINEAR", typeof(string));
            dt.Columns.Add("STATION_IFC", typeof(string));
            dt.Columns.Add("MM_BK", typeof(string));
            dt.Columns.Add("WALL_BK", typeof(string));
            dt.Columns.Add("PIPE_BK", typeof(string));
            dt.Columns.Add("HEAT_BK", typeof(string));
            dt.Columns.Add("COATING_BK", typeof(string));
            dt.Columns.Add("GRADE_BK", typeof(string));

            dt.Columns.Add("MM_AHD", typeof(string));
            dt.Columns.Add("WALL_AHD", typeof(string));
            dt.Columns.Add("PIPE_AHD", typeof(string));
            dt.Columns.Add("HEAT_AHD", typeof(string));
            dt.Columns.Add("COATING_AHD", typeof(string));
            dt.Columns.Add("GRADE_AHD", typeof(string));

            dt.Columns.Add("NG", typeof(string));
            dt.Columns.Add("NG_NORTHING", typeof(string));
            dt.Columns.Add("NG_EASTING", typeof(string));
            dt.Columns.Add("NG_ELEVATION", typeof(string));
            dt.Columns.Add("COVER", typeof(string));
            dt.Columns.Add("LOCATION", typeof(string));
            dt.Columns.Add("FILENAME", typeof(string));
            dt.Columns.Add("H_ANGLE", typeof(string));
            dt.Columns.Add("V_ANGLE", typeof(string));
            dt.Columns.Add("CROSSING_NAME", typeof(string));


            return dt;
        }


        public static void Create_weldmap_od_table()
        {

            string col1 = "MMID";
            string col2 = "PIPEID";
            string col3 = "HEAT";
            string col4 = "XRAY_START";
            string col5 = "XRAY_END";
            string col6 = "COATING";
            string col7 = "WALL";
            string col8 = "STA_START";
            string col9 = "STA_END";
            string col10 = "PNT_START";
            string col11 = "PNT_END";

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        List<string> List1 = new List<string>();
                        List<string> List2 = new List<string>();
                        List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();


                        List1.Add(col1);
                        List2.Add(col1);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col2);
                        List2.Add(col2);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col3);
                        List2.Add(col3);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col4);
                        List2.Add(col4);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col5);
                        List2.Add(col5);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col6);
                        List2.Add(col6);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col7);
                        List2.Add(col7);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col8);
                        List2.Add(col8);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col9);
                        List2.Add(col9);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col10);
                        List2.Add(col10);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        List1.Add(col11);
                        List2.Add(col11);
                        List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("WGEN_PIPE_REP", "Generated by WGEN", List1, List2, List3);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static System.Data.DataTable Creaza_all_points_datatable_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("PNT", typeof(string));//1
            dt.Columns.Add("NORTHING", typeof(string));//2
            dt.Columns.Add("EASTING", typeof(string));//3
            dt.Columns.Add("ELEVATION", typeof(string));//4
            dt.Columns.Add("FEATURE CODE", typeof(string));//5
            dt.Columns.Add("FILENAME", typeof(string));//6
            dt.Columns.Add("LOCATION", typeof(string));//7
            dt.Columns.Add("NOTES", typeof(string));//8
            dt.Columns.Add("DESCRIPTION", typeof(string));//9
            dt.Columns.Add("MISC1", typeof(string));//10
            dt.Columns.Add("MISC2", typeof(string));//11
            dt.Columns.Add("MISC3", typeof(string));//12
            dt.Columns.Add("MISC4", typeof(string));//13
            dt.Columns.Add("MISC5", typeof(string));//14
            dt.Columns.Add("MISC6", typeof(string));//15
            dt.Columns.Add("MISC7", typeof(string));//16

            return dt;
        }

        public static System.Data.DataTable creaza_error_export_table(System.Data.DataTable dt_err, string tabname)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            if (dt_err != null && dt_err.Rows.Count > 0)
            {
                dt1.Columns.Add("Point(MMid)", typeof(string));
                dt1.Columns.Add("TAB Name", typeof(string));
                dt1.Columns.Add("Error", typeof(string));
                dt1.Columns.Add("Value", typeof(string));
                dt1.Columns.Add("Excel", typeof(string));

                for (int i = 0; i < dt_err.Rows.Count; ++i)
                {
                    dt1.Rows.Add();
                    if (dt_err.Columns.Contains("Point") == true)
                    {
                        if (dt_err.Rows[i]["Point"] != DBNull.Value)
                        {
                            dt1.Rows[i]["Point(MMid)"] = Convert.ToString(dt_err.Rows[i]["Point"]);

                        }
                    }

                    dt1.Rows[i]["TAB Name"] = tabname;
                    if (dt_err.Rows[i]["Error"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Error"] = Convert.ToString(dt_err.Rows[i]["Error"]);

                    }
                    if (dt_err.Rows[i]["Value"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Value"] = Convert.ToString(dt_err.Rows[i]["Value"]);

                    }
                    if (dt_err.Rows[i]["Excel"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Excel"] = Convert.ToString(dt_err.Rows[i]["Excel"]);

                    }
                }
            }


            return dt1;
        }

        public static System.Data.DataTable creaza_error_export_table_for_pipe_manifest(System.Data.DataTable dt_err, string tabname)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            if (dt_err != null && dt_err.Rows.Count > 0)
            {
                dt1.Columns.Add("Point(MMid)", typeof(string));
                dt1.Columns.Add("TAB Name", typeof(string));
                dt1.Columns.Add("Error", typeof(string));
                dt1.Columns.Add("Value1", typeof(string));
                dt1.Columns.Add("Value2", typeof(string));
                dt1.Columns.Add("Excel", typeof(string));

                for (int i = 0; i < dt_err.Rows.Count; ++i)
                {
                    dt1.Rows.Add();
                    if (dt_err.Columns.Contains("Point") == true)
                    {
                        if (dt_err.Rows[i]["Point"] != DBNull.Value)
                        {
                            dt1.Rows[i]["Point(MMid)"] = Convert.ToString(dt_err.Rows[i]["Point"]);

                        }
                    }

                    dt1.Rows[i]["TAB Name"] = tabname;
                    if (dt_err.Rows[i]["Error"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Error"] = Convert.ToString(dt_err.Rows[i]["Error"]);

                    }
                    if (dt_err.Rows[i]["Value1"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Value1"] = Convert.ToString(dt_err.Rows[i]["Value1"]);

                    }
                    if (dt_err.Rows[i]["Value2"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Value2"] = Convert.ToString(dt_err.Rows[i]["Value2"]);

                    }
                    if (dt_err.Rows[i]["Excel"] != DBNull.Value)
                    {
                        dt1.Rows[i]["Excel"] = Convert.ToString(dt_err.Rows[i]["Excel"]);

                    }
                }
            }


            return dt1;
        }

        public static void Transfer_datatable_to_new_excel_spreadsheet_formated_general(System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "General";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                }
            }
        }

        public static void Transfer_datatable_to_new_excel_spreadsheet_formated_for_ground_tally(System.Data.DataTable dt1)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Get_NEW_worksheet_from_Excel();
                    W1.Name = "GROUND_TALLY";
                    W1.Columns["A:C"].NumberFormat = "@";
                    W1.Columns["D"].NumberFormat = "0.0";
                    W1.Columns["F"].NumberFormat = "@";
                    W1.Columns["G"].NumberFormat = "0";
                    W1.Columns["H"].NumberFormat = "@";

                    W1.Columns["N:BA"].NumberFormat = "@";
                    W1.Columns["A:K"].ColumnWidth = 15.11;



                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;
                }
            }
        }

        static public System.Data.DataTable build_data_table_from_excel_based_on_columns
      (
          System.Data.DataTable dt1,
          Microsoft.Office.Interop.Excel.Worksheet W1, int start_row,
          string col1, string colpt1, string col2, string colpt2, string col3, string colpt3, string col4, string colpt4,
          string col5, string colpt5, string col6, string colpt6, string col7, string colpt7, string col8, string colpt8,
          string col9, string colpt9, string col10, string colpt10, string col11, string colpt11,
          string col12, string colpt12, string col13, string colpt13, string col14, string colpt14,
          string col15, string colpt15, string col16, string colpt16, string col17, string colpt17, string col18, string colpt18,
          string col19, string colpt19, string col20, string colpt20, string col21, string colpt21,
          string col22, string colpt22, string col23, string colpt23, string col24, string colpt24,
          string col25, string colpt25, string col26, string colpt26, string col27, string colpt27
      )
        {
            if (W1 == null) return dt1;


            object[,] values1 = new object[300000, 1];
            object[,] values2 = new object[300000, 1];
            object[,] values3 = new object[300000, 1];
            object[,] values4 = new object[300000, 1];
            object[,] values5 = new object[300000, 1];
            object[,] values6 = new object[300000, 1];
            object[,] values7 = new object[300000, 1];
            object[,] values8 = new object[300000, 1];
            object[,] values9 = new object[300000, 1];
            object[,] values10 = new object[300000, 1];
            object[,] values11 = new object[300000, 1];
            object[,] values12 = new object[300000, 1];
            object[,] values13 = new object[300000, 1];
            object[,] values14 = new object[300000, 1];
            object[,] values15 = new object[300000, 1];
            object[,] values16 = new object[300000, 1];
            object[,] values17 = new object[300000, 1];
            object[,] values18 = new object[300000, 1];
            object[,] values19 = new object[300000, 1];
            object[,] values20 = new object[300000, 1];
            object[,] values21 = new object[300000, 1];
            object[,] values22 = new object[300000, 1];
            object[,] values23 = new object[300000, 1];
            object[,] values24 = new object[300000, 1];
            object[,] values25 = new object[300000, 1];
            object[,] values26 = new object[300000, 1];
            object[,] values27 = new object[300000, 1];

            #region 1
            if (colpt1 != "")
            {
                string adresa = colpt1 + start_row.ToString() + ":" + colpt1 + "30000";
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[adresa];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    if (Valoare1 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values1.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values1[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col1] = Valoare;
                }
            }
            #endregion

            #region 2
            if (colpt2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colpt2 + Convert.ToString(start_row) + ":" + colpt2 + "30000"];

                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (Valoare2 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }

                        }

                    }
                    else
                    {
                        i = values2.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values2[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col2] = Valoare;
                }
            }
            #endregion

            #region 3
            if (colpt3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colpt3 + Convert.ToString(start_row) + ":" + colpt3 + "30000"];

                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }

                        }

                    }
                    else
                    {
                        i = values3.Length + 1;
                    }

                }



                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values3[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col3] = Valoare;
                }
            }
            #endregion

            #region 4
            if (colpt4 != "")
            {
                Microsoft.Office.Interop.Excel.Range range4 = W1.Range[colpt4 + Convert.ToString(start_row) + ":" + colpt4 + "30000"];

                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (Valoare4 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                        }

                    }
                    else
                    {
                        i = values4.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values4[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col4] = Valoare;
                }
            }
            #endregion

            #region 5
            if (colpt5 != "")
            {
                Microsoft.Office.Interop.Excel.Range range5 = W1.Range[colpt5 + Convert.ToString(start_row) + ":" + colpt5 + "30000"];

                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (Valoare5 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                        }

                    }
                    else
                    {
                        i = values5.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values5[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col5] = Valoare;
                }
            }
            #endregion

            #region 6
            if (colpt6 != "")
            {
                Microsoft.Office.Interop.Excel.Range range6 = W1.Range[colpt6 + Convert.ToString(start_row) + ":" + colpt6 + "30000"];

                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (Valoare6 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();


                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                        }

                    }
                    else
                    {
                        i = values6.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values6[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col6] = Valoare;
                }
            }
            #endregion

            #region 7
            if (colpt7 != "")
            {
                Microsoft.Office.Interop.Excel.Range range7 = W1.Range[colpt7 + Convert.ToString(start_row) + ":" + colpt7 + "30000"];

                values7 = range7.Value2;

                for (int i = 1; i <= values7.Length; ++i)
                {
                    object Valoare7 = values7[i, 1];
                    if (Valoare7 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                        }

                    }
                    else
                    {
                        i = values7.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values7[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col7] = Valoare;
                }
            }
            #endregion

            #region 8
            if (colpt8 != "")
            {
                Microsoft.Office.Interop.Excel.Range range8 = W1.Range[colpt8 + Convert.ToString(start_row) + ":" + colpt8 + "30000"];

                values8 = range8.Value2;

                for (int i = 1; i <= values8.Length; ++i)
                {
                    object Valoare8 = values8[i, 1];
                    if (Valoare8 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }

                        }

                    }
                    else
                    {
                        i = values8.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values8[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col8] = Valoare;
                }
            }
            #endregion

            #region 9
            if (colpt9 != "")
            {
                Microsoft.Office.Interop.Excel.Range range9 = W1.Range[colpt9 + Convert.ToString(start_row) + ":" + colpt9 + "30000"];

                values9 = range9.Value2;

                for (int i = 1; i <= values9.Length; ++i)
                {
                    object Valoare9 = values9[i, 1];
                    if (Valoare9 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }
                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }
                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }
                        }
                    }
                    else
                    {
                        i = values9.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values9[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col9] = Valoare;
                }
            }
            #endregion

            #region 10
            if (colpt10 != "")
            {
                Microsoft.Office.Interop.Excel.Range range10 = W1.Range[colpt10 + Convert.ToString(start_row) + ":" + colpt10 + "30000"];

                values10 = range10.Value2;

                for (int i = 1; i <= values10.Length; ++i)
                {
                    object Valoare10 = values10[i, 1];
                    if (Valoare10 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                        }

                    }
                    else
                    {
                        i = values10.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values10[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col10] = Valoare;
                }
            }
            #endregion

            #region 11
            if (colpt11 != "")
            {
                Microsoft.Office.Interop.Excel.Range range11 = W1.Range[colpt11 + Convert.ToString(start_row) + ":" + colpt11 + "30000"];

                values11 = range11.Value2;

                for (int i = 1; i <= values11.Length; ++i)
                {
                    object Valoare11 = values11[i, 1];
                    if (Valoare11 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }
                        }

                    }
                    else
                    {
                        i = values11.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values11[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col11] = Valoare;
                }
            }
            #endregion

            #region 12

            if (colpt12 != "")
            {
                Microsoft.Office.Interop.Excel.Range range12 = W1.Range[colpt12 + Convert.ToString(start_row) + ":" + colpt12 + "30000"];

                values12 = range12.Value2;

                for (int i = 1; i <= values12.Length; ++i)
                {
                    object Valoare12 = values12[i, 1];
                    if (Valoare12 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }
                        }

                    }
                    else
                    {
                        i = values12.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values12[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col12] = Valoare;
                }
            }
            #endregion

            #region 13
            if (colpt13 != "")
            {
                Microsoft.Office.Interop.Excel.Range range13 = W1.Range[colpt13 + Convert.ToString(start_row) + ":" + colpt13 + "30000"];

                values13 = range13.Value2;

                for (int i = 1; i <= values13.Length; ++i)
                {
                    object Valoare13 = values13[i, 1];
                    if (Valoare13 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }
                        }

                    }
                    else
                    {
                        i = values13.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values13[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col13] = Valoare;
                }
            }
            #endregion

            #region 14
            if (colpt14 != "")
            {
                Microsoft.Office.Interop.Excel.Range range14 = W1.Range[colpt14 + Convert.ToString(start_row) + ":" + colpt14 + "30000"];

                values14 = range14.Value2;

                for (int i = 1; i <= values14.Length; ++i)
                {
                    object Valoare14 = values14[i, 1];
                    if (Valoare14 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }
                        }

                    }
                    else
                    {
                        i = values14.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values14[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col14] = Valoare;
                }
            }
            #endregion

            #region 15
            if (colpt15 != "")
            {
                Microsoft.Office.Interop.Excel.Range range15 = W1.Range[colpt15 + Convert.ToString(start_row) + ":" + colpt15 + "30000"];

                values15 = range15.Value2;

                for (int i = 1; i <= values15.Length; ++i)
                {
                    object Valoare15 = values15[i, 1];
                    if (Valoare15 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }
                        }

                    }
                    else
                    {
                        i = values15.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values15[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col15] = Valoare;
                }
            }
            #endregion

            #region 16
            if (colpt16 != "")
            {
                Microsoft.Office.Interop.Excel.Range range16 = W1.Range[colpt16 + Convert.ToString(start_row) + ":" + colpt16 + "30000"];

                values16 = range16.Value2;

                for (int i = 1; i <= values16.Length; ++i)
                {
                    object Valoare16 = values16[i, 1];
                    if (Valoare16 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }
                        }

                    }
                    else
                    {
                        i = values16.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values16[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col16] = Valoare;
                }
            }
            #endregion

            #region 17
            if (colpt17 != "")
            {
                Microsoft.Office.Interop.Excel.Range range17 = W1.Range[colpt17 + Convert.ToString(start_row) + ":" + colpt17 + "30000"];

                values17 = range17.Value2;

                for (int i = 1; i <= values17.Length; ++i)
                {
                    object Valoare17 = values17[i, 1];
                    if (Valoare17 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }
                        }

                    }
                    else
                    {
                        i = values17.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values17[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col17] = Valoare;
                }
            }

            #endregion

            #region 18
            if (colpt18 != "")
            {
                Microsoft.Office.Interop.Excel.Range range18 = W1.Range[colpt18 + Convert.ToString(start_row) + ":" + colpt18 + "30000"];

                values18 = range18.Value2;

                for (int i = 1; i <= values18.Length; ++i)
                {
                    object Valoare18 = values18[i, 1];
                    if (Valoare18 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                        }

                    }
                    else
                    {
                        i = values18.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values18[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col18] = Valoare;
                }
            }
            #endregion

            #region 19
            if (colpt19 != "")
            {
                Microsoft.Office.Interop.Excel.Range range19 = W1.Range[colpt19 + Convert.ToString(start_row) + ":" + colpt19 + "30000"];

                values19 = range19.Value2;

                for (int i = 1; i <= values19.Length; ++i)
                {
                    object Valoare19 = values19[i, 1];
                    if (Valoare19 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                        }

                    }
                    else
                    {
                        i = values19.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values19[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col19] = Valoare;
                }
            }
            #endregion

            #region 20
            if (colpt20 != "")
            {
                Microsoft.Office.Interop.Excel.Range range20 = W1.Range[colpt20 + Convert.ToString(start_row) + ":" + colpt20 + "30000"];

                values20 = range20.Value2;

                for (int i = 1; i <= values20.Length; ++i)
                {
                    object Valoare20 = values20[i, 1];
                    if (Valoare20 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                        }

                    }
                    else
                    {
                        i = values20.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values20[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col20] = Valoare;
                }
            }
            #endregion

            #region 21
            if (colpt21 != "")
            {
                Microsoft.Office.Interop.Excel.Range range21 = W1.Range[colpt21 + Convert.ToString(start_row) + ":" + colpt21 + "30000"];

                values21 = range21.Value2;

                for (int i = 1; i <= values21.Length; ++i)
                {
                    object Valoare21 = values21[i, 1];
                    if (Valoare21 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                        }

                    }
                    else
                    {
                        i = values21.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values21[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col21] = Valoare;
                }
            }
            #endregion

            #region 22
            if (colpt22 != "")
            {
                Microsoft.Office.Interop.Excel.Range range22 = W1.Range[colpt22 + Convert.ToString(start_row) + ":" + colpt22 + "30000"];

                values22 = range22.Value2;

                for (int i = 1; i <= values22.Length; ++i)
                {
                    object Valoare22 = values22[i, 1];
                    if (Valoare22 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }

                        }

                    }
                    else
                    {
                        i = values22.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values22[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col22] = Valoare;
                }
            }
            #endregion

            #region 23
            if (colpt23 != "")
            {
                Microsoft.Office.Interop.Excel.Range range23 = W1.Range[colpt23 + Convert.ToString(start_row) + ":" + colpt23 + "30000"];

                values23 = range23.Value2;

                for (int i = 1; i <= values23.Length; ++i)
                {
                    object Valoare23 = values23[i, 1];
                    if (Valoare23 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                        }

                    }
                    else
                    {
                        i = values23.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values23[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col23] = Valoare;
                }
            }

            #endregion

            #region 24

            if (colpt24 != "")
            {
                Microsoft.Office.Interop.Excel.Range range24 = W1.Range[colpt24 + Convert.ToString(start_row) + ":" + colpt24 + "30000"];

                values24 = range24.Value2;

                for (int i = 1; i <= values24.Length; ++i)
                {
                    object Valoare24 = values24[i, 1];
                    if (Valoare24 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                        }

                    }
                    else
                    {
                        i = values24.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values24[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col24] = Valoare;
                }
            }

            #endregion

            #region 25
            if (colpt25 != "")
            {
                Microsoft.Office.Interop.Excel.Range range25 = W1.Range[colpt25 + Convert.ToString(start_row) + ":" + colpt25 + "30000"];

                values25 = range25.Value2;

                for (int i = 1; i <= values25.Length; ++i)
                {
                    object Valoare25 = values25[i, 1];
                    if (Valoare25 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }

                        }

                    }
                    else
                    {
                        i = values25.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values25[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col25] = Valoare;
                }
            }
            #endregion

            #region 26
            if (colpt26 != "")
            {
                Microsoft.Office.Interop.Excel.Range range26 = W1.Range[colpt26 + Convert.ToString(start_row) + ":" + colpt26 + "30000"];

                values26 = range26.Value2;

                for (int i = 1; i <= values26.Length; ++i)
                {
                    object Valoare26 = values26[i, 1];
                    if (Valoare26 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }


                            if (colpt25 != "")
                            {
                                object Valoare25 = values25[i + 1, 1];
                                if (Valoare25 == null) Valoare25 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col25] = Valoare25;
                            }

                        }

                    }
                    else
                    {
                        i = values26.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values26[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col26] = Valoare;
                }
            }
            #endregion

            #region 27
            if (colpt27 != "")
            {
                Microsoft.Office.Interop.Excel.Range range27 = W1.Range[colpt27 + Convert.ToString(start_row) + ":" + colpt27 + "30000"];

                values27 = range27.Value2;

                for (int i = 1; i <= values27.Length; ++i)
                {
                    object Valoare27 = values27[i, 1];
                    if (Valoare27 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }

                            if (colpt11 != "")
                            {
                                object Valoare11 = values11[i + 1, 1];
                                if (Valoare11 == null) Valoare11 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col11] = Valoare11;
                            }

                            if (colpt12 != "")
                            {
                                object Valoare12 = values12[i + 1, 1];
                                if (Valoare12 == null) Valoare12 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col12] = Valoare12;
                            }

                            if (colpt13 != "")
                            {
                                object Valoare13 = values13[i + 1, 1];
                                if (Valoare13 == null) Valoare13 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col13] = Valoare13;
                            }

                            if (colpt14 != "")
                            {
                                object Valoare14 = values14[i + 1, 1];
                                if (Valoare14 == null) Valoare14 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col14] = Valoare14;
                            }

                            if (colpt15 != "")
                            {
                                object Valoare15 = values15[i + 1, 1];
                                if (Valoare15 == null) Valoare15 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col15] = Valoare15;
                            }

                            if (colpt16 != "")
                            {
                                object Valoare16 = values16[i + 1, 1];
                                if (Valoare16 == null) Valoare16 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col16] = Valoare16;
                            }

                            if (colpt17 != "")
                            {
                                object Valoare17 = values17[i + 1, 1];
                                if (Valoare17 == null) Valoare17 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col17] = Valoare17;
                            }

                            if (colpt18 != "")
                            {
                                object Valoare18 = values18[i + 1, 1];
                                if (Valoare18 == null) Valoare18 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col18] = Valoare18;
                            }

                            if (colpt19 != "")
                            {
                                object Valoare19 = values19[i + 1, 1];
                                if (Valoare19 == null) Valoare19 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col19] = Valoare19;
                            }

                            if (colpt20 != "")
                            {
                                object Valoare20 = values20[i + 1, 1];
                                if (Valoare20 == null) Valoare20 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col20] = Valoare20;
                            }

                            if (colpt21 != "")
                            {
                                object Valoare21 = values21[i + 1, 1];
                                if (Valoare21 == null) Valoare21 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col21] = Valoare21;
                            }


                            if (colpt22 != "")
                            {
                                object Valoare22 = values22[i + 1, 1];
                                if (Valoare22 == null) Valoare22 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col22] = Valoare22;
                            }

                            if (colpt23 != "")
                            {
                                object Valoare23 = values23[i + 1, 1];
                                if (Valoare23 == null) Valoare23 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col23] = Valoare23;
                            }

                            if (colpt24 != "")
                            {
                                object Valoare24 = values24[i + 1, 1];
                                if (Valoare24 == null) Valoare24 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col24] = Valoare24;
                            }


                            if (colpt25 != "")
                            {
                                object Valoare25 = values25[i + 1, 1];
                                if (Valoare25 == null) Valoare25 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col25] = Valoare25;
                            }

                            if (colpt26 != "")
                            {
                                object Valoare26 = values26[i + 1, 1];
                                if (Valoare26 == null) Valoare26 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col26] = Valoare26;
                            }

                        }

                    }
                    else
                    {
                        i = values27.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values27[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col27] = Valoare;
                }
            }
            #endregion

            return dt1;
        }

        static public System.Data.DataTable build_data_table_from_excel_based_on_11_columns_for_pipe_tally
(
  System.Data.DataTable dt1,
  Microsoft.Office.Interop.Excel.Worksheet W1, int start_row,
  string col1, string colpt1, string col2, string colpt2, string col3, string colpt3, string col4, string colpt4,
  string col5, string colpt5, string col6, string colpt6, string col7, string colpt7, string col8, string colpt8,
  string col9, string colpt9, string col10, string colpt10, string col11, string colpt11
)
        {
            if (W1 == null) return dt1;


            object[,] values1 = new object[30000, 1];
            object[,] values2 = new object[30000, 1];
            object[,] values3 = new object[30000, 1];
            object[,] values4 = new object[30000, 1];
            object[,] values5 = new object[30000, 1];
            object[,] values6 = new object[30000, 1];
            object[,] values7 = new object[30000, 1];
            object[,] values8 = new object[30000, 1];
            object[,] values9 = new object[30000, 1];
            object[,] values10 = new object[30000, 1];
            object[,] values11 = new object[30000, 1];


            #region 1
            if (colpt1 != "")
            {
                string adresa = colpt1 + start_row.ToString() + ":" + colpt1 + "30000";
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[adresa];
                values1 = range1.Value2;

                for (int i = 1; i <= values1.Length; ++i)
                {
                    object Valoare1 = values1[i, 1];
                    if (Valoare1 != null)
                    {
                        dt1.Rows.Add();
                    }
                    else
                    {
                        i = values1.Length + 1;
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values1[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col1] = Valoare;
                }
            }
            #endregion

            #region 2
            if (colpt2 != "")
            {
                Microsoft.Office.Interop.Excel.Range range2 = W1.Range[colpt2 + Convert.ToString(start_row) + ":" + colpt2 + "30000"];

                values2 = range2.Value2;

                for (int i = 1; i <= values2.Length; ++i)
                {
                    object Valoare2 = values2[i, 1];
                    if (Valoare2 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }

                        }

                    }
                    else
                    {
                        i = values2.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values2[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col2] = Valoare;
                }
            }
            #endregion

            #region 3
            if (colpt3 != "")
            {
                Microsoft.Office.Interop.Excel.Range range3 = W1.Range[colpt3 + Convert.ToString(start_row) + ":" + colpt3 + "30000"];

                values3 = range3.Value2;

                for (int i = 1; i <= values3.Length; ++i)
                {
                    object Valoare3 = values3[i, 1];
                    if (Valoare3 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }

                        }

                    }
                    else
                    {
                        i = values3.Length + 1;
                    }

                }



                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values3[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col3] = Valoare;
                }
            }
            #endregion

            #region 4
            if (colpt4 != "")
            {
                Microsoft.Office.Interop.Excel.Range range4 = W1.Range[colpt4 + Convert.ToString(start_row) + ":" + colpt4 + "30000"];

                values4 = range4.Value2;

                for (int i = 1; i <= values4.Length; ++i)
                {
                    object Valoare4 = values4[i, 1];
                    if (Valoare4 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                        }

                    }
                    else
                    {
                        i = values4.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values4[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col4] = Valoare;
                }
            }
            #endregion

            #region 5
            if (colpt5 != "")
            {
                Microsoft.Office.Interop.Excel.Range range5 = W1.Range[colpt5 + Convert.ToString(start_row) + ":" + colpt5 + "30000"];

                values5 = range5.Value2;

                for (int i = 1; i <= values5.Length; ++i)
                {
                    object Valoare5 = values5[i, 1];
                    if (Valoare5 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                        }

                    }
                    else
                    {
                        i = values5.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values5[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col5] = Valoare;
                }
            }
            #endregion

            #region 6
            if (colpt6 != "")
            {
                Microsoft.Office.Interop.Excel.Range range6 = W1.Range[colpt6 + Convert.ToString(start_row) + ":" + colpt6 + "30000"];

                values6 = range6.Value2;

                for (int i = 1; i <= values6.Length; ++i)
                {
                    object Valoare6 = values6[i, 1];
                    if (Valoare6 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();


                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                        }

                    }
                    else
                    {
                        i = values6.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values6[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col6] = Valoare;
                }
            }
            #endregion

            #region 7
            if (colpt7 != "")
            {
                Microsoft.Office.Interop.Excel.Range range7 = W1.Range[colpt7 + Convert.ToString(start_row) + ":" + colpt7 + "30000"];

                values7 = range7.Value2;

                for (int i = 1; i <= values7.Length; ++i)
                {
                    object Valoare7 = values7[i, 1];
                    if (Valoare7 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                        }

                    }
                    else
                    {
                        i = values7.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values7[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col7] = Valoare;
                }
            }
            #endregion

            #region 8
            if (colpt8 != "")
            {
                Microsoft.Office.Interop.Excel.Range range8 = W1.Range[colpt8 + Convert.ToString(start_row) + ":" + colpt8 + "30000"];

                values8 = range8.Value2;

                for (int i = 1; i <= values8.Length; ++i)
                {
                    object Valoare8 = values8[i, 1];
                    if (Valoare8 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }

                        }

                    }
                    else
                    {
                        i = values8.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values8[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col8] = Valoare;
                }
            }
            #endregion

            #region 9
            if (colpt9 != "")
            {
                Microsoft.Office.Interop.Excel.Range range9 = W1.Range[colpt9 + Convert.ToString(start_row) + ":" + colpt9 + "30000"];

                values9 = range9.Value2;

                for (int i = 1; i <= values9.Length; ++i)
                {
                    object Valoare9 = values9[i, 1];
                    if (Valoare9 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();
                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }
                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }
                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }
                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }
                        }
                    }
                    else
                    {
                        i = values9.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values9[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col9] = Valoare;
                }
            }
            #endregion

            #region 10
            if (colpt10 != "")
            {
                Microsoft.Office.Interop.Excel.Range range10 = W1.Range[colpt10 + Convert.ToString(start_row) + ":" + colpt10 + "30000"];

                values10 = range10.Value2;

                for (int i = 1; i <= values10.Length; ++i)
                {
                    object Valoare10 = values10[i, 1];
                    if (Valoare10 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                        }

                    }
                    else
                    {
                        i = values10.Length + 1;
                    }

                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values10[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col10] = Valoare;
                }
            }
            #endregion

            #region 11
            if (colpt11 != "")
            {
                Microsoft.Office.Interop.Excel.Range range11 = W1.Range[colpt11 + Convert.ToString(start_row) + ":" + colpt11 + "30000"];

                values11 = range11.Value2;

                for (int i = 1; i <= values11.Length; ++i)
                {
                    object Valoare11 = values11[i, 1];
                    if (Valoare11 != null)
                    {
                        if (i > dt1.Rows.Count)
                        {
                            dt1.Rows.Add();

                            if (colpt1 != "")
                            {
                                object Valoare1 = values1[i + 1, 1];
                                if (Valoare1 == null) Valoare1 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col1] = Valoare1;
                            }
                            if (colpt2 != "")
                            {
                                object Valoare2 = values2[i + 1, 1];
                                if (Valoare2 == null) Valoare2 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col2] = Valoare2;
                            }
                            if (colpt3 != "")
                            {
                                object Valoare3 = values3[i + 1, 1];
                                if (Valoare3 == null) Valoare3 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col3] = Valoare3;
                            }
                            if (colpt4 != "")
                            {
                                object Valoare4 = values4[i + 1, 1];
                                if (Valoare4 == null) Valoare4 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col4] = Valoare4;
                            }

                            if (colpt5 != "")
                            {
                                object Valoare5 = values5[i + 1, 1];
                                if (Valoare5 == null) Valoare5 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col5] = Valoare5;
                            }

                            if (colpt6 != "")
                            {
                                object Valoare6 = values6[i + 1, 1];
                                if (Valoare6 == null) Valoare6 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col6] = Valoare6;
                            }


                            if (colpt7 != "")
                            {
                                object Valoare7 = values7[i + 1, 1];
                                if (Valoare7 == null) Valoare7 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col7] = Valoare7;
                            }
                            if (colpt8 != "")
                            {
                                object Valoare8 = values8[i + 1, 1];
                                if (Valoare8 == null) Valoare8 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col8] = Valoare8;
                            }

                            if (colpt9 != "")
                            {
                                object Valoare9 = values9[i + 1, 1];
                                if (Valoare9 == null) Valoare9 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col9] = Valoare9;
                            }

                            if (colpt10 != "")
                            {
                                object Valoare10 = values10[i + 1, 1];
                                if (Valoare10 == null) Valoare10 = DBNull.Value;
                                dt1.Rows[dt1.Rows.Count - 1][col10] = Valoare10;
                            }
                        }

                    }
                    else
                    {
                        i = values11.Length + 1;
                    }



                }

                for (int i = 0; i < dt1.Rows.Count; ++i)
                {
                    object Valoare = values11[i + 1, 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt1.Rows[i][col11] = Valoare;
                }
            }
            #endregion


            #region extra columns

            object[,] values_extra = new object[79, 1];
            Microsoft.Office.Interop.Excel.Range range_extra = W1.Range["A1:CA1"];
            values_extra = range_extra.Value2;

            int nr_col0 = dt1.Columns.Count;
            int maxRows = dt1.Rows.Count;

            dt1.Columns.Add("_");
            dt1.Columns.Add("__");

            for (int i = 1; i <= values_extra.Length; ++i)
            {
                object Valoare1 = values_extra[1, i];
                if (Valoare1 != null)
                {
                    string nume_col = Convert.ToString(Valoare1);
                    if (dt1.Columns.Contains(nume_col) == true)
                    {
                        nume_col = nume_col + "_original";
                    }

                    dt1.Columns.Add(nume_col, typeof(string));

                }
                else
                {
                    i = values_extra.Length;
                }

            }

            int maxCols = dt1.Columns.Count - nr_col0 - 2;
            string col_end = get_excel_column_letter(maxCols);

            Range range_orig = W1.Range["A2:" + col_end + Convert.ToString(maxRows + 1)];



            object[,] values_orig = new object[maxRows, maxCols];
            values_orig = range_orig.Value2;


            for (int i = 0; i < maxRows; ++i)
            {
                for (int j = 0; j < maxCols; ++j)
                {
                    dt1.Rows[i][j + nr_col0 + 2] = values_orig[i + 1, j + 1];
                }
            }

            #endregion

            return dt1;
        }

        public static string get_excel_column_letter(int intCol)
        {

            string columnString = "";
            decimal columnNumber = intCol;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }

        public static Microsoft.Office.Interop.Excel.Worksheet Transfer_weldmap_datatable_to_new_excel_spreadsheet_formated_general_and_colored(System.Data.DataTable dt1, System.Data.DataTable dt2)
        {
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                   W1 = Get_NEW_worksheet_from_Excel();
                    W1.Cells.NumberFormat = "General";
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int i = 0; i < dt1.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName;
                    }
                    range1.Value2 = values1;

                    W1.Range["A:A"].Font.Bold = true;
                    W1.Range["I:I"].Font.Bold = true;
                    W1.Range["O:O"].Font.Bold = true;
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        for (int i = 0; i < maxRows; ++i)
                        {
                            if (dt2.Rows[i][0] != DBNull.Value)
                            {
                                int color1 = Convert.ToInt32(dt2.Rows[i][0]);
                                W1.Rows[2 + i].Interior.Color = color1;
                            }
                            if (dt2.Rows[i][1] != DBNull.Value)
                            {
                                int cidx = Convert.ToInt32(dt2.Rows[i][1]);
                                W1.Rows[2 + i].Interior.ColorIndex = cidx;
                            }
                            if (dt2.Rows[i][2] != DBNull.Value)
                            {
                                int thc = Convert.ToInt32(dt2.Rows[i][2]);
                                W1.Rows[2 + i].Interior.ThemeColor = thc;
                            }
                            if (dt2.Rows[i][3] != DBNull.Value)
                            {
                                double tint1 = Convert.ToDouble(dt2.Rows[i][3]);
                                W1.Rows[2 + i].Interior.PatternTintAndShade = tint1;
                            }
                        }
                    }
                    W1.Range["A:A"].ColumnWidth = 7.29;
                    W1.Range["B:D"].ColumnWidth = 0;
                    W1.Range["E:E"].ColumnWidth = 22.86;
                    W1.Range["F:F"].ColumnWidth = 58;
                    W1.Range["G:H"].ColumnWidth = 21.57;
                    W1.Range["G:H"].NumberFormat = "0+00";
                    W1.Name = "WELD_MAP";
                }
            }
            return W1;
        }


        public static void create_backup(string fisier1)
        {
            if (System.IO.File.Exists(fisier1) == false)
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Microsoft.Office.Interop.Excel.Workbook Workbook1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return;
                for (int k = 1; k <= Excel1.Workbooks.Count; ++k)
                {
                    Workbook1 = Excel1.Workbooks[k];
                    string wn = Workbook1.Name;

                    if (wn.ToUpper() == fisier1.ToUpper())
                    {
                        fisier1 = Workbook1.FullName;
                        k = Excel1.Workbooks.Count;
                    }
                }
            }


            if (System.IO.File.Exists(fisier1) == true)
            {
                string Director1 = System.IO.Path.GetDirectoryName(fisier1);
                if (Director1.Substring(Director1.Length - 1, 1) != "\\")
                {
                    Director1 = Director1 + "\\";
                }

                string name1 = System.IO.Path.GetFileNameWithoutExtension(fisier1);
                string backup1 = Director1 + "~Archive";
                if (System.IO.Directory.Exists(backup1) == false)
                {
                    System.IO.Directory.CreateDirectory(backup1);
                }
                string backup2 = "C:\\Users\\Public\\" + "~Archive";
                if (System.IO.Directory.Exists(backup2) == false && Environment.UserName.ToUpper() == "POP70694")
                {
                    System.IO.Directory.CreateDirectory(backup2);
                }

                string new_name = name1 + "-[" + System.DateTime.Now.Year.ToString() + "_" + System.DateTime.Now.Month.ToString() + "_" + System.DateTime.Now.Day.ToString() +
                    "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString() + "_" + System.DateTime.Now.Second.ToString() +
                  "]-" + Environment.UserName.ToUpper() + ".xlsx";
                backup1 = backup1 + "\\" + new_name;
                backup2 = backup2 + "\\" + new_name;

                System.IO.File.Copy(fisier1, backup1);

                if (Environment.UserName.ToUpper() == "POP70694")
                {
                    System.IO.File.Copy(fisier1, backup2);
                }
            }
        }

    }



}
