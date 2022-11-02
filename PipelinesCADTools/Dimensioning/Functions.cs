using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Windows.Forms;


namespace Dimensioning
{
    class Functions
    {
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

        static public String verify_layername_from_combobox_different_database(Database Database1, System.Windows.Forms.ComboBox Combo_layername)
        {

            string Layer_name = "0";
            if (Combo_layername.Text != "")
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable Layer_table = Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as LayerTable;

                    if (Layer_table.Has(Combo_layername.Text) == true)
                    {
                        Layer_name = Combo_layername.Text;
                    }
                    Trans1.Dispose();
                }




            }


            return Layer_name;
        }

        static public void Creaza_layer_with_linetype(string Layername, short Culoare, string Ltypename, bool Plot)
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
                        LinetypeTable Linetype_table = Trans1.GetObject(ThisDrawing.Database.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                        if (LayerTable1.Has(Layername) == true && Linetype_table.Has(Ltypename) == true)
                        {
                            LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                            new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                            new_layer.LinetypeObjectId = Linetype_table[Ltypename];
                            new_layer.IsPlottable = Plot;

                        }

                        if (LayerTable1.Has(Layername) == false && Linetype_table.Has(Ltypename) == true)
                        {
                            LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                            new_layer.Name = Layername;
                            new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                            new_layer.LinetypeObjectId = Linetype_table[Ltypename];
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


        static public double GET_Bearing_rad(Double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
        }

        static public string Get_chainage_feet_from_double(Double Numar, int Nr_dec)
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
                String3 = String2.Substring(0, String2.Length - 2 - Nr_dec - Punct) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));
            }
            else
            {
                if (String2.Length - Nr_dec - Punct == 1) String3 = "0+0" + String2;
                if (String2.Length - Nr_dec - Punct == 2) String3 = "0+" + String2;
                if (String2.Length - Nr_dec - Punct == 3) String3 = String2.Substring(0, 1) + "+" + String2.Substring(String2.Length - (2 + Nr_dec + Punct));

            }
            return String_minus + String3;

        }

        static public string Get_String_Rounded(Double Numar, int Nr_dec)
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

        static public string Get_String_Rounded_with_thousand_sep(Double Numar, int Nr_dec)
        {

            String String1, String2, Zero, zero1;
            Zero = "";
            zero1 = "";
            string Comma = ",";

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
                            String4 = String4 + "0";
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

        static public string Get_DMS(Double Numar)
        {
            int Degree1 = Convert.ToInt32(Math.Floor(Numar));
            int Minutes1 = Convert.ToInt32(Math.Floor((Numar - Convert.ToDouble(Degree1)) * 60));

            double REST = Convert.ToDouble(Degree1) + Convert.ToDouble(Minutes1) / 60;
            double Sec1 = (Numar - REST) * 3600;

            int Seconds1 = Convert.ToInt32(Math.Round(Sec1, 0));

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

            String D = Degree1.ToString();
            String M = Minutes1.ToString();
            String S = Seconds1.ToString();

            if (M.Length == 1)
            {
                M = "0" + M;
            }

            if (S.Length == 1)
            {
                S = "0" + S;
            }

            int ii = 176;
            char S1 = (char)ii;

            int ij = 34;
            char q1 = (char)ij;

            return D + S1 + M + "'" + S + q1;
        }

        static public string Get_0DMS(Double Numar)
        {
            int Degree1 = Convert.ToInt32(Math.Floor(Numar));
            int Minutes1 = Convert.ToInt32(Math.Floor((Numar - Convert.ToDouble(Degree1)) * 60));

            double REST = Convert.ToDouble(Degree1) + Convert.ToDouble(Minutes1) / 60;
            double Sec1 = (Numar - REST) * 3600;

            int Seconds1 = Convert.ToInt32(Math.Round(Sec1, 0));

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

            String D = Degree1.ToString();
            if (Math.Abs(Degree1) < 10) D = "0" + D;

            String M = Minutes1.ToString();
            String S = Seconds1.ToString();

            if (M.Length == 1)
            {
                M = "0" + M;
            }

            if (S.Length == 1)
            {
                S = "0" + S;
            }

            int ii = 176;
            char S1 = (char)ii;

            int ij = 34;
            char q1 = (char)ij;

            return D + S1 + M + "'" + S + q1;
        }


        static public string Get_Quadrant_bearing(double Radian1, bool spatiu_dupa_val )
        {

            string space1 = " ";
            if (spatiu_dupa_val == false)
            {
                space1 = "";
            }
            string Prefix1 = "N"+space1;
            string Suffix1 = space1+"E";
            double Quadrant1 = Math.PI / 2 - Radian1;
            


            if (Radian1 > Math.PI / 2 & Radian1 <= Math.PI)
            {
                Quadrant1 = Radian1 - Math.PI / 2;
                Suffix1 = space1 + "W";
            }
            if (Radian1 > Math.PI & Radian1 <= 3 * Math.PI / 2)
            {
                Quadrant1 = 3 * Math.PI / 2 - Radian1;
                Prefix1 = "S" + space1;
                Suffix1 = space1 + "W";
            }
            if (Radian1 > 3 * Math.PI / 2)
            {
                Quadrant1 = Radian1 - 3 * Math.PI / 2;
                Prefix1 = "S" + space1;
                Suffix1 = space1 + "E";
            }
            return Prefix1 + Get_DMS(Quadrant1 * 180 / Math.PI) + Suffix1;
        }

        static public double calculate_rotatie_text(double Radian1)
        {
            Double Rot_t = Radian1;
            if (Radian1 > Math.PI / 2 & Radian1 <= Math.PI)
            {
                Rot_t = Radian1 + Math.PI;
            }
            if (Radian1 > Math.PI & Radian1 <= 3 * Math.PI / 2)
            {
                Rot_t = Radian1 + Math.PI;
            }
            return Rot_t;
        }

        static public void Incarca_existing_Blocks_with_attributes_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    Combo_blockname.Items.Clear();
                    foreach (ObjectId Block_id in BlockTable_data1)
                    {
                        BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                        if (Block1.HasAttributeDefinitions == true)
                        {
                            Combo_blockname.Items.Add(Block1.Name);
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

        static public BlockReference InsertBlock_with_multiple_atributes(string Nume_fisier, string NumeBlock, Point3d Insertion_point, double Scale_xyz, string Layer1,
             System.Collections.Specialized.StringCollection Colectie_nume_atribute, System.Collections.Specialized.StringCollection Colectie_valori_atribute)
        {

            BlockReference Block1 = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    if (BlockTable1.Has(NumeBlock) == false)
                    {
                        if (System.IO.File.Exists(Nume_fisier) == true)
                        {
                            using (Database Database2 = new Database(false, false))
                            {
                                Database2.ReadDwgFile(Nume_fisier, System.IO.FileShare.Read, true, null);
                                ThisDrawing.Database.Insert(NumeBlock, Database2, false);
                            }
                        }


                    }

                    Trans1.Commit();
                }

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                    if (BlockTable1.Has(NumeBlock) == true)
                    {

                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTR = (BlockTableRecord)Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForRead);

                        Block1 = new BlockReference(Insertion_point, BTR.ObjectId);
                        Block1.Layer = Layer1;
                        Block1.ScaleFactors = new Autodesk.AutoCAD.Geometry.Scale3d(Scale_xyz, Scale_xyz, Scale_xyz);
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
                                    if (Attref.Tag == Tag1)
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
            }
            return Block1;
        }

        static public String get_block_name(BlockReference Block1)
        {
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    BlockTableRecord Btr = null;
                    if (Block1.IsDynamicBlock == true)
                    {

                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                    else
                    {
                        Btr = (BlockTableRecord)Trans1.GetObject(Block1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                        return Btr.Name;
                    }
                }
            }
            catch (System.Exception ex)
            {
                return "";
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

        public static void Populate_object_data_table(Autodesk.Gis.Map.ObjectData.Tables Tables1, ObjectId oB1, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {

                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), oB1, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
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
                                Tabla1.AddRecord(rec, oB1);
                            }
                        }
                    }
                }
            }
        }


        static public void Incarca_existing_layers_to_combobox(System.Windows.Forms.ComboBox Combo_layer)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                Combo_layer.Items.Clear();

                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));


                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;


                    }
                }

                System.Data.DataTable dt2 = Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    Combo_layer.Items.Add(dt2.Rows[i][0].ToString());
                }
                Combo_layer.SelectedIndex = 0;
                Trans1.Dispose();
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
                            Double param_on_1 = Curba1.GetParameterAtPoint(Pt1);
                            Double param_on_2 = Curba2.GetParameterAtPoint(Pt1);


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


        public static System.Data.DataTable Creaza_config_datatable_structure()
        {





            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();


            Lista1.Add("Client");
            Lista1.Add("Layer");
            Lista1.Add("Layer_color");
            Lista1.Add("Ltscale");
            Lista1.Add("Line_def1");
            Lista1.Add("Line_def2");
            Lista1.Add("extension");

            Lista2.Add(typeof(string));
            Lista2.Add(typeof(string));
            Lista2.Add(typeof(int));
            Lista2.Add(typeof(double));
            Lista2.Add(typeof(string));
            Lista2.Add(typeof(string));
            Lista2.Add(typeof(double));

            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }

        public static System.Data.DataTable Build_Data_table_CONFIG_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {


            System.Data.DataTable Data_table_config = Creaza_config_datatable_structure();
            int NrR = 0;
            int NrC = Data_table_config.Columns.Count;



            bool is_data = false;

            for (int i = Start_row; i < 30000; ++i)
            {
                if (i == Start_row)
                {
                    if (W1.Range["A" + i.ToString()].Value2 == null)
                    {
                        MessageBox.Show("no data found in the CLIENT config file");
                        return Data_table_config;
                    }
                }

                if (W1.Range["A" + i.ToString()].Value2 == null)
                {
                    NrR = i - Start_row;
                    i = 31000;
                }
                else
                {
                    Data_table_config.Rows.Add();
                    is_data = true;
                }
            }


            if (is_data == true)
            {

                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < Data_table_config.Rows.Count; ++i)
                {
                    for (int j = 0; j < Data_table_config.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;

                        Data_table_config.Rows[i][j] = Valoare;
                    }
                }
            }

            Data_table_config.Columns.Add("Line_name", typeof(string));
            if (is_data == true)
            {
                for (int i = 0; i < Data_table_config.Rows.Count; ++i)
                {
                    if (Data_table_config.Rows[i][4] != DBNull.Value)
                    {
                        string descr1 = Data_table_config.Rows[i][4].ToString();
                        if (descr1.Contains(",") == true)
                        {
                            int pos1 = descr1.IndexOf(",");

                            Data_table_config.Rows[i][7] = descr1.Substring(1, pos1 - 1);
                        }
                    }
                }
            }

            return Data_table_config;


        }
        public static double Round5(double x)
        {
            return Math.Round(x / 5, MidpointRounding.AwayFromZero) * 5;
        }

    }
}





namespace Autodesk.AutoCAD.DatabaseServices
{
    public static class ViewportExtensionMethods
    {
        public static Matrix3d GetModelToPaperTransform(this Viewport vport)
        {
            if (vport.PerspectiveOn)
                throw new NotSupportedException("Perspective views not supported");
            Point3d center = new Point3d(vport.ViewCenter.X, vport.ViewCenter.Y, 0.0);
            Vector3d v = new Vector3d(vport.CenterPoint.X - center.X, vport.CenterPoint.Y - center.Y, 0.0);
            return Matrix3d.Displacement(v)
               * Matrix3d.Scaling(vport.CustomScale, center)
               * Matrix3d.Rotation(vport.TwistAngle, Vector3d.ZAxis, Point3d.Origin)
               * Matrix3d.WorldToPlane(new Plane(vport.ViewTarget, vport.ViewDirection));
        }

        /// <summary>
        /// Indicates if the viewport's extents contains the given point.
        /// </summary>
        /// <remarks>
        /// Does not support clipped viewports. Used primarily for trivial
        /// rejection of candidate points prior to containment testing against
        /// non-rectangular viewort clipping boundaries.
        /// </remarks>
        /// <param name="vport">The Viewport</param>
        /// <param name="point">The model space world coordinate to test</param>
        /// <param name="checkFrontBackClipping">A boolean indicating if
        /// front/back clipping should be tested</param>
        /// <returns>True if the point is within the viewport and is between
        /// the front/back clipping planes, if either/both are enabled</returns>

        public static bool ContainsPoint(this Viewport vport, Point3d point, bool checkFrontBackClipping = true)
        {
            point = point.TransformBy(GetModelToPaperTransform(vport));
            Extents3d vpextents = vport.GeometricExtents;
            if (!vpextents.Contains(point, true))
                return false;
            if (checkFrontBackClipping)
            {
                if (vport.BackClipOn && point.Z < (vport.BackClipDistance * vport.CustomScale))
                    return false;
                if (vport.FrontClipOn && point.Z > (vport.FrontClipDistance * vport.CustomScale))
                    return false;
            }
            return true;
        }
    }
}

namespace Autodesk.AutoCAD.Geometry
{
    // excerpts from GeometryExtensions

    public static class GeometryExtensions
    {
        /// <summary>
        /// Indicates if the given bounding box contains the given Point3d
        /// </summary>
        /// <param name="extents">An Extents3d representing the bounding box</param>
        /// <param name="point">The Point3d to test for containment</param>
        /// <param name="project">If true, the base of the bounding box and the point
        /// are projected into the XY plane, and the result indicates if the projected
        /// point is contained within the projected rectangle</param>
        /// <returns>true if the extents contains the point, or its projection into
        /// the XY plane contains the point projected into the XY plane</returns>

        public static bool Contains(this Extents3d extents, Point3d point, bool project = false)
        {
            if (project)
                return Contains2d(extents.MinPoint, extents.MaxPoint, point);
            else
                return Contains(extents.MinPoint, extents.MaxPoint, point);
        }

        static bool Contains2d(Point3d min, Point3d max, Point3d p)
        {
            return !(p.X < min.X || p.X > max.X || p.Y < min.Y || p.Y > max.Y);
        }

        static bool Contains(Point3d min, Point3d max, Point3d p)
        {
            return !(p.X < min.X || p.X > max.X || p.Y < min.Y || p.Y > max.Y || p.Z < min.Z || p.Z > max.Z);
        }



    }
}