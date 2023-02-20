using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Management;
using System.Windows.Forms;

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

        public static bool isSECURE()
        {

            string number_drive = GetHDDSerialNumber("C");

            switch (number_drive)
            {
                case "123":
                    return true;
                default:
                    try
                    {
                        string UserDNS = Environment.GetEnvironmentVariable("USERDNSDOMAIN");
                        if (UserDNS.ToUpper() == "HMMG.CC" | UserDNS.ToLower() == "mottmac.group.int")
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    catch (System.Exception)
                    {
                        return false;
                    }
            }
        }


        public static string GetHDDSerialNumber(string drive)
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

        public static void make_first_layout_active(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    LayoutManager1.CurrentLayout = Layout0.LayoutName;
                    return;
                }

            }

        }


        public static void make_layout_active(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, int no = 1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == no)
                {
                    LayoutManager1.CurrentLayout = Layout0.LayoutName;
                    return;
                }

            }

        }

        public static string get_layout_name(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, int no = 1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == no)
                {
                    return Layout0.LayoutName;
                }

            }

            return "";

        }

        public static int no_of_paperspace_layouts(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            int no_of_layouts = 1;

            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder > 1)
                {
                    ++no_of_layouts;
                }

            }

            return no_of_layouts;
        }


        public static void Color_border_range_inside(Microsoft.Office.Interop.Excel.Range range1, int cid)
        {

            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range1.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternNone;
            range1.Interior.TintAndShade = 0;
            range1.Interior.PatternTintAndShade = 0;
            if (cid != 0)
            {
                range1.Interior.ColorIndex = cid;
            }

        }

        public static void Color_border_range_outside(Microsoft.Office.Interop.Excel.Range range1, int cid)
        {


            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
            range1.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternNone;
            range1.Interior.TintAndShade = 0;
            range1.Interior.PatternTintAndShade = 0;
            if (cid != 0)
            {
                range1.Interior.ColorIndex = cid;
            }
        }



        static public bool Get_if_workbook_is_open_in_Excel(string filename)
        {

            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                if (Excel1 == null) return false;

                foreach (Workbook wbk in Excel1.Workbooks)
                {
                    if (wbk.Name.ToLower() == filename.ToLower())
                    {
                        return true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                return false;
            }
            return false;
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
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return 0;
            }



        }


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


        public static void create_backup(string fisier1)
        {
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
        public static void Transfer_to_worksheet_Data_table(Worksheet W1, System.Data.DataTable Data_table, int Start_row, string format_cell)
        {


            W1.Columns["A:XX"].Delete();

            if (Data_table != null)
            {
                if (Data_table.Rows.Count > 0)
                {
                    int NrR = Data_table.Rows.Count;
                    int NrC = Data_table.Columns.Count;


                    Object[,] values = new object[NrR + 1, NrC];
                    for (int i = 0; i < NrR; ++i)
                    {
                        for (int j = 0; j < NrC; ++j)
                        {
                            if (Data_table.Rows[i][j] != DBNull.Value)
                            {
                                values[i + 1, j] = Data_table.Rows[i][j];
                            }
                        }
                    }
                    for (int j = 0; j < NrC; ++j)
                    {
                        values[0, j] = Data_table.Columns[j].ColumnName;
                    }

                    Microsoft.Office.Interop.Excel.Range range0 = W1.Range[W1.Columns[1], W1.Columns[NrC]];
                    range0.ClearContents();
                    range0.UnMerge();
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row, NrC]];
                    range1.Cells.NumberFormat = format_cell;
                    range1.Value2 = values;
                    Color_border_range_inside(range1, 0);
                }
            }
        }

        public static void delete_entities_with_OD(string layer_name, string od_table_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                {
                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    foreach (ObjectId id1 in BTrecord)
                    {
                        Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                        if (ent1 != null)
                        {
                            if (ent1.Layer == layer_name)
                            {
                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                bool delete1 = false;
                                if (Tables1.IsTableDefined(od_table_name) == true)
                                {
                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[od_table_name];
                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                    {
                                        if (Records1.Count > 0)
                                        {
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                            {
                                                if (delete1 == false)
                                                {
                                                    for (int i = 0; i < Record1.Count; ++i)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                        string Nume_field = Field_def1.Name;
                                                        string Valoare1 = Record1[i].StrValue;
                                                        if (Nume_field == "SegmentName")
                                                        {
                                                            string segment1 = "";
                                                            if (Valoare1 == segment1)
                                                            {
                                                                delete1 = true;
                                                                i = Record1.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (delete1 == true)
                                    {
                                        ent1.UpgradeOpen();
                                        ent1.Erase();
                                    }
                                }
                            }
                        }
                    }
                    Trans1.Commit();
                }
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

        public static Autodesk.AutoCAD.DatabaseServices.BlockTableRecord get_modelspace(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;
            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)Database1.BlockTableId.GetObject(OpenMode.ForRead);
            return (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(BlockTable_data1[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        }

        public static Autodesk.AutoCAD.DatabaseServices.BlockTableRecord get_first_layout_as_paperspace(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {

            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordPS = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    return (BlockTableRecord)Trans1.GetObject(Layout0.BlockTableRecordId, OpenMode.ForRead);
                }

            }
            return BTrecordPS;


        }

        public static Autodesk.AutoCAD.DatabaseServices.BlockTableRecord get_layout_as_paperspace(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, int no = 1)
        {

            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordPS = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == no)
                {
                    return (BlockTableRecord)Trans1.GetObject(Layout0.BlockTableRecordId, OpenMode.ForRead);
                }

            }
            return BTrecordPS;


        }

        public static Layout get_first_layout(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            Layout Layout0 = null;
            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout0 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout0.TabOrder == 1)
                {
                    return Layout0;
                }

            }
            return Layout0;
        }


        public static Layout add_new_layout(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, string layoutname)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            DBDictionary Layoutdict = (DBDictionary)Trans1.GetObject(Database1.LayoutDictionaryId, OpenMode.ForRead);

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

            foreach (DBDictionaryEntry entry in Layoutdict)
            {
                Layout Layout1 = (Layout)Trans1.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                if (Layout1.LayoutName == layoutname)
                {
                    MessageBox.Show("the layout named " + layoutname + " is already in the drawing\r\nOperation aborted!");
                    return null;
                }

            }

            ObjectId id1 = LayoutManager1.CreateLayout(layoutname);
            Layout Layout0 = (Layout)Trans1.GetObject(id1, OpenMode.ForRead);
            LayoutManager1.CurrentLayout = layoutname;

            return Layout0;
        }


        public static void make_ms_active(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1)
        {
            HostApplicationServices.WorkingDatabase = Database1;

            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
            LayoutManager1.CurrentLayout = "Model";


        }

        public static Viewport Create_viewport(Point3d MSpoint, Point3d PSpoint, double Width, double Height, double Scale, double Twist_rad)
        {
            Viewport Viewport1 = new Viewport();

            Viewport1.SetDatabaseDefaults();
            Viewport1.CenterPoint = PSpoint;
            Viewport1.Height = Height;
            Viewport1.Width = Width;
            Viewport1.ViewDirection = Autodesk.AutoCAD.Geometry.Vector3d.ZAxis;
            Viewport1.ViewTarget = MSpoint;
            Viewport1.ViewCenter = Autodesk.AutoCAD.Geometry.Point2d.Origin;
            Viewport1.TwistAngle = Twist_rad;
            Viewport1.CustomScale = Scale;
            Viewport1.Locked = true;


            return Viewport1;
        }

        static public void Creaza_layer_on_database(Database Database1, string Layername, short Culoare, bool Plot)
        {
            try
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.LayerTable LayerTable1;
                    LayerTable1 = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(Database1.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    if (LayerTable1.Has(Layername) == true)
                    {
                        LayerTable1.UpgradeOpen();
                        LayerTableRecord new_layer = Trans1.GetObject(LayerTable1[Layername], OpenMode.ForWrite) as LayerTableRecord;
                        if (new_layer != null)
                        {
                            new_layer.IsPlottable = Plot;
                            Trans1.Commit();
                        }
                    }

                    if (LayerTable1.Has(Layername) == false)
                    {
                        LayerTable1.UpgradeOpen();
                        LayerTableRecord new_layer = new Autodesk.AutoCAD.DatabaseServices.LayerTableRecord();
                        new_layer.Name = Layername;
                        new_layer.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Culoare);
                        new_layer.IsPlottable = Plot;
                        LayerTable1.Add(new_layer);
                        Trans1.AddNewlyCreatedDBObject(new_layer, true);
                        Trans1.Commit();
                    }


                    Trans1.Dispose();

                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        public static void Transfer_datatable_to_new_excel_spreadsheet(System.Data.DataTable dt1, string sheetname = "Sheet1")
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

                    for (int j = 0; j < maxCols; ++j)
                    {
                        string column_letter = get_excel_column_letter(j + 1);
                        if (dt1.Columns[j].DataType == typeof(double))
                        {
                            W1.Range[column_letter + ":" + column_letter].NumberFormat = "0.000";
                        }
                        else if (dt1.Columns[j].DataType == typeof(int))
                        {
                            W1.Range[column_letter + ":" + column_letter].NumberFormat = "0";
                        }
                        else if (dt1.Columns[j].DataType == typeof(string))
                        {
                            W1.Range[column_letter + ":" + column_letter].NumberFormat = "@";
                        }
                    }

                    range1.Value2 = values1;
                    W1.Name = sheetname;

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

        static public String get_block_name_another_database(BlockReference Block1, Database database2)
        {
            try
            {

                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = database2.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)database2.BlockTableId.GetObject(OpenMode.ForRead);
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

        public static void zoom_to_Point(Point3d pt, double zoom_delta_distance)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - zoom_delta_distance, pt.Y - zoom_delta_distance, 0);
                            Point3d maxx = new Point3d(pt.X + zoom_delta_distance, pt.Y + zoom_delta_distance, 0);

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

        static public void Incarca_existing_layers_to_combobox(System.Windows.Forms.ComboBox Combo_layer)
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
                            string Name_of_layer = Layer1.Name;
                            if (Name_of_layer.Contains("|") == false & Name_of_layer.Contains("$") == false)
                            {
                                Array.Resize(ref Array1, idx1);
                                Array1[idx1 - 1] = Name_of_layer;
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
                            string nume1 = Block1.Name;
                            if (nume1.Contains("*") == false) Combo_blockname.Items.Add(nume1);
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

        static public string get_block_name(BlockReference Block1)
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
            catch (System.Exception)
            {
                return "";
            }

        }


        static public void Update_Attrib_block_values(BlockReference Block1, System.Collections.Specialized.StringCollection Col_name, System.Collections.Specialized.StringCollection Col_value)
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);



                    if (Block1.AttributeCollection.Count > 0 & Col_name != null & Col_value != null)
                    {

                        if (Col_name.Count == Col_value.Count)
                        {
                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;

                            foreach (ObjectId ID1 in attColl)
                            {
                                DBObject ent = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                if (ent is AttributeReference)
                                {
                                    AttributeReference attref = (AttributeReference)ent;
                                    attref.UpgradeOpen();

                                    if (Col_name.Contains(attref.Tag) == true)
                                    {
                                        int index1 = Col_name.IndexOf(attref.Tag);
                                        attref.TextString = Col_value[index1];
                                    }
                                }

                            }

                        }
                    }
                    Trans1.Commit();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }



        }
        static public System.Data.DataTable Read_block_attributes_and_values(BlockReference Block1)
        {
            System.Data.DataTable Table1 = new System.Data.DataTable();
            Table1.Columns.Add("ATTRIB", typeof(string));
            Table1.Columns.Add("VALUE", typeof(string));


            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                    if (Block1.AttributeCollection.Count > 0)
                    {
                        Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = Block1.AttributeCollection;

                        foreach (ObjectId ID1 in attColl)
                        {
                            DBObject ent = Trans1.GetObject(ID1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            if (ent is AttributeReference)
                            {
                                AttributeReference attref = (AttributeReference)ent;
                                Table1.Rows.Add();
                                Table1.Rows[Table1.Rows.Count - 1]["ATTRIB"] = attref.Tag;
                                if (attref.IsMTextAttribute == false)
                                {
                                    Table1.Rows[Table1.Rows.Count - 1]["VALUE"] = attref.TextString;
                                }
                                if (attref.IsMTextAttribute == true)
                                {
                                    Table1.Rows[Table1.Rows.Count - 1]["VALUE"] = attref.MTextAttribute.Contents;
                                }
                            }

                        }

                    }
                    Trans1.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return Table1;


        }


        public static Polyline Build_2dpoly_from_3d(Polyline3d Poly3D)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    Polyline Poly2D = new Polyline();
                    int Index1 = 0;
                    if (Poly3D.Length > 0)
                    {

                        double last_param = Poly3D.EndParam;

                        for (int i = 0; i <= last_param; ++i)
                        {
                            try
                            {
                                Point3d pt1 = Poly3D.GetPointAtParameter(i).TransformBy(curent_ucs_matrix);
                                Poly2D.AddVertexAt(Index1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                ++Index1;

                            }
                            catch (System.Exception)
                            {

                            }
                        }
                    }
                    return Poly2D;
                }
            }







        }

        static public void add_OD_fieds_to_combobox(string table_name, ComboBox Combobox1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    if (Tables1.IsTableDefined(table_name) == true)
                    {
                        Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[table_name];
                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                        for (int i = 0; i < Field_defs1.Count; ++i)
                        {
                            Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[i];

                            if (Combobox1.Items.Contains(fielddef1.Name) == false)
                            {
                                Combobox1.Items.Add(fielddef1.Name);
                            }
                        }
                    }
                    else
                    {
                        Combobox1.Items.Clear();
                    }
                    Trans1.Commit();
                }
            }
        }

        public static void creaza_anno_scales(Database Database2)
        {
            List<string> lista_anno_names = new List<string>();
            List<double> lista_anno_ps = new List<double>();
            lista_anno_names.Add("1:10");
            lista_anno_ps.Add(10);
            lista_anno_names.Add("1:20");
            lista_anno_ps.Add(20);
            lista_anno_names.Add("1:30");
            lista_anno_ps.Add(30);
            lista_anno_names.Add("1:40");
            lista_anno_ps.Add(40);
            lista_anno_names.Add("1:50");
            lista_anno_ps.Add(50);
            lista_anno_names.Add("1:60");
            lista_anno_ps.Add(60);
            lista_anno_names.Add("1:100");
            lista_anno_ps.Add(100);
            lista_anno_names.Add("1:200");
            lista_anno_ps.Add(200);
            lista_anno_names.Add("1:300");
            lista_anno_ps.Add(300);
            lista_anno_names.Add("1:400");
            lista_anno_ps.Add(400);
            lista_anno_names.Add("1:500");
            lista_anno_ps.Add(500);
            lista_anno_names.Add("1:600");
            lista_anno_ps.Add(600);


            var ocm = Database2.ObjectContextManager;
            var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");

            for (int i = 0; i < lista_anno_names.Count; i++)
            {
                AnnotationScale ano1 = new AnnotationScale();
                ano1.Name = lista_anno_names[i];
                ano1.PaperUnits = 1;
                ano1.DrawingUnits = lista_anno_ps[i];

                if (occ.HasContext(ano1.Name) == false)
                {
                    occ.AddContext(ano1);

                }
            }
        }
        static public void set_block_visibility(BlockReference BR, String visibility_name)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {


                    if (prop.PropertyName == "Visibility1" && !prop.ReadOnly)
                    {
                        object[] existing = prop.GetAllowedValues();
                        bool found = false;

                        foreach (object ob in existing)
                        {
                            if (ob.ToString() == visibility_name)
                            {
                                found = true;
                            }
                        }

                        if (found == true)
                        {
                            if (prop.Value.ToString() != visibility_name)
                            {
                                prop.Value = visibility_name;
                            }
                        }
                    }



                }
                return;
            }
        }
    }
}
