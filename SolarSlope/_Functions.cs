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
using Autodesk.Civil.DatabaseServices;
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

        static string get_W1_name()
        {
            return Environment.UserName + " " + DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at " + DateTime.Now.Hour + "h" + DateTime.Now.Minute + "m";
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


        static public Worksheet Create_a_new_worksheet_from_excel_by_name(string filename, string SheetName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Workbook Workbook1;
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

                        Worksheet W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W1.Name = SheetName;
                        return W1;
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

        public static void Transfer_datatable_to_existing_excel_spreadsheet_by_name(System.Data.DataTable dt1, string filename, string sheetname, bool hiden_sheet)
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

                            W1.Cells.ClearContents();
                            W1.Cells.ClearFormats();
                            W1.Cells.NumberFormat = "General";

                            if (dt1 != null && dt1.Rows.Count > 0)
                            {

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
                                    W1.Cells[1, i + 1].value2 = dt1.Columns[i].ColumnName.ToUpper();
                                }
                                range1.Value2 = values1;
                                if (hiden_sheet == false) W1.Visible = XlSheetVisibility.xlSheetHidden;
                                return;

                            }
                        }
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


        public static System.Data.DataTable Sort_data_table_desc(System.Data.DataTable Datatable1, string Column1)
        {
            System.Data.DataTable Data_table_temp = new System.Data.DataTable();
            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    if (Datatable1.Columns.Contains(Column1) == true)
                    {
                        System.Data.DataView DataView1 = new System.Data.DataView(Datatable1);
                        DataView1.Sort = Column1 + " DESC";
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



        public static Worksheet Transfer_datatable_to_new_excel_spreadsheet_formated_general(System.Data.DataTable dt1)
        {
            Worksheet W1 = null;
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
                    string new_name = Environment.UserName.ToUpper() + "_" + System.DateTime.Now.Year.ToString() + "_" + System.DateTime.Now.Month.ToString() + "_" + System.DateTime.Now.Day.ToString() +
                                    "_" + System.DateTime.Now.Hour.ToString() + "_" + System.DateTime.Now.Minute.ToString() ;
                    W1.Name = new_name;
                }
            }

            return W1;

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


        public static int get_excel_column_index(string col1)
        {
            int retVal = 0;
            string col = col1.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
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



        public static Microsoft.Office.Interop.Excel.Worksheet Transfer_datatable_to_new_excel_spreadsheet(System.Data.DataTable dt1, string sheetname = "Sheet1")
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

            return W1;
        }

        public static UDPString Find_udp_string(string udp_name1)
        {
            UDPString foundUDP = null;
            foreach (UDP udp1 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
            {
                if (udp1!= null && udp1.Name == udp_name1 && udp1.GetType().Name == "UDPString")
                {
                    foundUDP = udp1 as UDPString;
                }
            }
            return foundUDP;
        }



        public static UDPDouble Find_udp_double(string udp_name1)
        {
            UDPDouble foundUDP = null;
            foreach (UDP udp2 in Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument.PointUDPs)
            {
                if (udp2 != null && udp2.Name == udp_name1 && udp2.GetType().Name == "UDPDouble")
                {

                    foundUDP = udp2 as UDPDouble;
                }
            }
            return foundUDP;
        }



        public static Polyline Build_2d_poly_from_datatable(System.Data.DataTable dt_cl)
        {
            string Col_x = "X";
            string Col_y = "Y";

            Polyline Poly2D = new Polyline();

            int index1 = 0;

            for (int i = 0; i < dt_cl.Rows.Count; ++i)
            {
                double x = 0;
                double y = 0;

                if (dt_cl.Rows[i][Col_x] != DBNull.Value)
                {
                    x = (double)dt_cl.Rows[i][Col_x];
                    if (dt_cl.Rows[i][Col_y] != DBNull.Value)
                    {
                        y = (double)dt_cl.Rows[i][Col_y];

                        double bulge1 = 0;
                        if (dt_cl.Rows[i][0] != DBNull.Value)
                        {
                            string b1 = Convert.ToString(dt_cl.Rows[i][0]);
                            if (IsNumeric(b1) == true)
                            {
                                bulge1 = Convert.ToDouble(b1);
                            }

                        }


                        Poly2D.AddVertexAt(index1, new Point2d(x, y), bulge1, 0, 0);
                        Poly2D.Elevation = 0;

                        index1 = index1 + 1;
                    }
                }
            }

            return Poly2D;


        }

        public static MLeader creaza_mleader(Point3d pt_ins, string continut, double texth, double delta_x, double delta_y, double lgap, double dogl, double arrow, string layer1 = "0")
        {



            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            MLeader mleader1 = null;


            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {

                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                mleader1 = new MLeader();
                mleader1.SetDatabaseDefaults();
                mleader1.ContentType = ContentType.MTextContent;
                mleader1.LeaderLineType = LeaderType.StraightLeader;
                mleader1.Annotative = AnnotativeStates.False;

                MText mtext1 = new MText();
                mtext1.SetDatabaseDefaults();
                mtext1.Contents = continut;
                //mtext1.TextHeight = texth;
                mtext1.BackgroundFill = true;
                mtext1.UseBackgroundColor = true;
                mtext1.BackgroundScaleFactor = 1.2;
                mtext1.ColorIndex = 0;
                mleader1.MText = mtext1;

                int index1 = mleader1.AddLeader();
                int index2 = mleader1.AddLeaderLine(index1);
                mleader1.AddFirstVertex(index2, pt_ins);
                mleader1.AddLastVertex(index2, new Point3d(pt_ins.X + delta_x, pt_ins.Y + delta_y, pt_ins.Z));


                mleader1.TextHeight = texth;

                mleader1.LandingGap = lgap;
                mleader1.ArrowSize = arrow;
                mleader1.DoglegLength = dogl;
                mleader1.ColorIndex = 256;
                mleader1.Layer = layer1;
                BTrecord.AppendEntity(mleader1);
                Trans1.AddNewlyCreatedDBObject(mleader1, true);
                Trans1.Commit();




            }




            return mleader1;







        }

        static public string Get_chainage_from_double(double Numar, string units, int Nr_dec)
        {

            string String2, String3;
            String3 = "";
            string String_minus = "";

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
        static public void Load_existing_Blocks_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
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

                    List<string> lista1 = new List<string>();
                    foreach (ObjectId Block_id in BlockTable_data1)
                    {
                        BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                        if (Block1.Name.Contains("*") == false && Block1.Name.Contains("|") == false &&
                            Block1.Name.Contains("$") == false && Block1.IsFromExternalReference == false &&
                            Block1.IsFromOverlayReference == false &&
                            Block1.IsLayout == false)
                        {
                            lista1.Add(Block1.Name);
                        }
                    }
                    if (lista1.Count > 0)
                    {
                        Combo_blockname.Items.Add("");
                        lista1.Sort();
                        for (int i = 0; i < lista1.Count; ++i)
                        {
                            Combo_blockname.Items.Add(lista1[i]);
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

        static public List<string> Incarca_existing_Atributes_to_list(string BlockName)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable Block_table = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTable;

                List<string> Lista1 = new List<string>();

                if (BlockName != "" && Block_table != null)
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecordBlock = Trans1.GetObject(Block_table[BlockName], Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTableRecord;
                    if (BTrecordBlock != null)
                    {
                        foreach (ObjectId Id1 in BTrecordBlock)
                        {
                            Autodesk.AutoCAD.DatabaseServices.Entity ent = Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                            if (ent != null)
                            {
                                AttributeDefinition attDefinition1 = ent as AttributeDefinition;
                                if (attDefinition1 != null)
                                {
                                    Lista1.Add(attDefinition1.Tag);
                                }
                            }
                        }
                    }
                }
                Trans1.Dispose();
                return Lista1;
            }
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
                        Autodesk.AutoCAD.DatabaseServices.Entity Ent1 = Trans1.GetObject(BTR_enum.Current, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Entity;
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



    }

}

