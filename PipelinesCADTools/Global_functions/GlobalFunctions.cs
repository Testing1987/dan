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
    public class global_functions
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

        public static ObjectId GetObjectId(Database db, string handle)
        {
            try
            {
                return db.GetObjectId(false, new Handle(Convert.ToInt64(handle)), 0);
            }
            catch (System.Exception EX)
            {
                //MessageBox.Show(EX.Message + "\r\nObject ID not present in the drawing database");
                return ObjectId.Null;
            }

        }

        public static void Populate_object_data_table_from_handle_string(Autodesk.Gis.Map.ObjectData.Tables Tables1, string ObjId, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {
                    ObjectId oB1 = GetObjectId(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database, ObjId);
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
                                        string Valoare = "";
                                        if (List_value[i] != null)
                                        {
                                            Valoare = List_value[i].ToString();
                                        }

                                        Val.Assign(Valoare);
                                    }
                                    if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                    {
                                        double Valoare = 0;
                                        if (List_value[i] != null)
                                        {
                                            Valoare = Convert.ToDouble(List_value[i]);
                                        }

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

    }

}
