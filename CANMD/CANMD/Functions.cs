
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    class Functions
    {
        public static bool isSECURE()
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

                case "120E4B54":
                    return true;
                case "F6633173":
                    return true;
                case "40D6BDCB":
                    return true;
                case "18399D24":
                    return true;

                case "B63AD3F6":
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
                    catch (System.Exception ex)
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

        public static string get_initial()
        {
            if (is_dan_popescu() == true)
            {
                return "DP";
            }
            else
            {
                return Environment.UserName.ToUpper();
            }
        }


        static public void Load_opened_workbooks_to_combobox(ComboBox combo1)
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
                    //for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    //{
                    //    combo1.Items.Add("[" + Workbook1.Worksheets[i].name + "] - " + wn);
                    //}

                    combo1.Items.Add(wn);

                }
                if (combo1.Items.Count > 0) combo1.SelectedIndex = 0;

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static public int get_dropdown_width(ComboBox combo1)
        {
            Graphics g = combo1.CreateGraphics();
            SizeF size;
            int new_size = combo1.Width + 20;

            foreach (string item1 in combo1.Items)
            {
                size = g.MeasureString(item1, combo1.Font);

                if (size.Width > new_size)
                {
                    new_size = (int)size.Width;

                }
            }

            return new_size;
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

                return null;
            }



        }

        static public System.Data.DataTable build_data_table_from_excel(System.Data.DataTable dt1, Microsoft.Office.Interop.Excel.Worksheet W1, int start_row, int end_row, List<string> list_col, List<string> list_colxl)
        {
            if (W1 == null) return dt1;
            if (end_row - start_row < 0) return dt1;
            if (list_col.Count == 0) return dt1;
            if (list_col.Count != list_colxl.Count) return dt1;
            for (int i = 0; i < 1; ++i)
            {
                object[,] values_i = new object[end_row - start_row + 1, 1];
                if (list_colxl[i] != "")
                {
                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[list_colxl[i] + start_row.ToString() + ":" + list_colxl[i] + end_row.ToString()];
                    values_i = range1.Value2;
                    for (int k = 1; k <= values_i.Length; ++k)
                    {
                        object Valoare_k = values_i[k, 1];
                        if (Valoare_k != null && dt1.Columns[list_col[i]].DataType == typeof(double))
                        {
                            Valoare_k = Valoare_k.ToString().Replace("+", "");
                            if (IsNumeric(Valoare_k.ToString()) == true)
                            {
                                Valoare_k = Convert.ToDouble(Valoare_k);
                            }
                            else
                            {
                                Valoare_k = DBNull.Value;
                            }
                        }
                        if (Valoare_k == null) Valoare_k = DBNull.Value;
                        if (Valoare_k != DBNull.Value)
                        {
                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1][list_col[i]] = Valoare_k;
                            if (dt1.Columns.Contains("Excel") == true)
                            {
                                if (dt1.Rows[dt1.Rows.Count - 1]["Excel"] == DBNull.Value)
                                {
                                    dt1.Rows[dt1.Rows.Count - 1]["Excel"] = k;
                                }
                            }
                        }
                        else
                        {
                            k = values_i.Length + 1;
                        }
                    }
                }
            }

            if (list_col.Count > 1)
            {
                for (int i = 1; i < list_col.Count; ++i)
                {
                    object[,] values_i = new object[end_row - start_row + 1, 1];
                    if (list_colxl[i] != "")
                    {
                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[list_colxl[i] + start_row.ToString() + ":" + list_colxl[i] + end_row.ToString()];
                        values_i = range1.Value2;
                        for (int k = 1; k <= values_i.Length; ++k)
                        {
                            object Valoare_k = values_i[k, 1];
                            if (Valoare_k != null && dt1.Columns[list_col[i]].DataType == typeof(double))
                            {
                                Valoare_k = Valoare_k.ToString().Replace("+", "");
                                if (IsNumeric(Valoare_k.ToString()) == true)
                                {
                                    Valoare_k = Convert.ToDouble(Valoare_k);
                                }
                                else
                                {
                                    Valoare_k = DBNull.Value;
                                }
                            }
                            if (Valoare_k == null) Valoare_k = DBNull.Value;

                            if (k - 1 < dt1.Rows.Count)
                            {
                                dt1.Rows[k - 1][list_col[i]] = Valoare_k;
                            }

                        }
                    }
                }
            }
            return dt1;
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

        public static System.Data.DataTable Sort_data_table_2_columns(System.Data.DataTable Datatable1, string Column1, string Column2)
        {
            System.Data.DataTable Data_table_temp = new System.Data.DataTable();
            if (Datatable1 != null)
            {
                if (Datatable1.Rows.Count > 0)
                {
                    if (Datatable1.Columns.Contains(Column1) == true)
                    {
                        System.Data.DataView DataView1 = new System.Data.DataView(Datatable1);
                        DataView1.Sort = Column1 + " ASC, " + Column2 + " ASC";
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

        public static Worksheet Transfer_datatable_to_new_excel_spreadsheet(System.Data.DataTable dt1, string sheetname = "Sheet1", List<string> lista_col = null, List<double> lista_width = null, bool general = false)
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
                    if (general == false)
                    {
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
                    }
                    else
                    {
                        W1.Range["A:ZZ"].NumberFormat = "General";
                    }


                    range1.Value2 = values1;
                    W1.Name = sheetname;

                    if (lista_col != null && lista_col.Count == lista_width.Count)
                    {
                        for (int i = 0; i < lista_col.Count; ++i)
                        {
                            W1.Range[lista_col[i] + ":" + lista_col[i]].ColumnWidth = lista_width[i];
                        }
                    }



                }
            }
            return W1;
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

        public static void Round_data_table(System.Data.DataTable dt1, int round1)
        {

            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    for (int j = 0; j < dt1.Columns.Count; ++j)
                    {
                        if (dt1.Columns[j].DataType == typeof(double))
                        {
                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                if (dt1.Rows[i][j] != DBNull.Value)
                                {
                                    double val1 = Convert.ToDouble(dt1.Rows[i][j]);
                                    val1 = Math.Round(val1, round1);
                                    dt1.Rows[i][j] = val1;
                                }
                            }
                        }
                    }

                }
            }
        }
        public static Polyline Build_2d_poly_for_scanning(System.Data.DataTable dt_cl)
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
                        if (dt_cl.Columns.Contains("Bulge") == true)
                        {
                            bulge1 = Convert.ToDouble(dt_cl.Rows[i]["Bulge"]);
                        }


                        Poly2D.AddVertexAt(index1, new Point2d(x, y), bulge1, 0, 0);
                        Poly2D.Elevation = 0;

                        index1 = index1 + 1;
                    }
                }
            }

            return Poly2D;


        }

        public static Polyline3d Build_3d_poly_for_scanning(System.Data.DataTable dt_cl)
        {
            string Col_x = "X";
            string Col_y = "Y";
            string Col_z = "Z";

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            Polyline3d Poly3D = new Polyline3d();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    BTrecord.AppendEntity(Poly3D);
                    Trans1.AddNewlyCreatedDBObject(Poly3D, true);



                    Poly3D.SetDatabaseDefaults();

                    for (int i = 0; i < dt_cl.Rows.Count; ++i)
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        if (dt_cl.Rows[i][Col_x] != DBNull.Value)
                        {
                            x = (double)dt_cl.Rows[i][Col_x];
                        }

                        if (dt_cl.Rows[i][Col_y] != DBNull.Value)
                        {
                            y = (double)dt_cl.Rows[i][Col_y];
                        }

                        if (dt_cl.Rows[i][Col_z] != DBNull.Value)
                        {
                            z = (double)dt_cl.Rows[i][Col_z];
                        }


                        PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                        Poly3D.AppendVertex(Vertex_new);
                        Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                    }

                    Trans1.Commit();
                }
            }
            return Poly3D;

        }


        public static Polyline get_part_of_poly(Polyline poly0, double par1, double par2)
        {
            if (par1 > par2)
            {
                double t = par1;
                par1 = par2;
                par2 = t;
            }

            if (par2 > poly0.EndParam) par2 = poly0.EndParam;

            Polyline poly1 = new Polyline();
            int idx = 0;

            poly1.AddVertexAt(idx, new Point2d(poly0.GetPointAtParameter(par1).X, poly0.GetPointAtParameter(par1).Y), 0, 0, 0);
            ++idx;
            for (int i = 0; i < poly0.NumberOfVertices; ++i)
            {
                if (i > par1 && i < par2)
                {
                    poly1.AddVertexAt(idx, poly0.GetPoint2dAt(i), 0, 0, 0);
                    ++idx;
                }
            }

            poly1.AddVertexAt(idx, new Point2d(poly0.GetPointAtParameter(par2).X, poly0.GetPointAtParameter(par2).Y), 0, 0, 0);

            Point2d pt_prev = new Point2d();
            double bear_prev = -1000000;
            for (int i = poly1.NumberOfVertices - 1; i >= 0; --i)
            {
                Point2d pt1 = poly1.GetPoint2dAt(i);
                double d1 = Math.Round(Math.Pow(Math.Pow(pt_prev.X - pt1.X, 2) + Math.Pow(pt_prev.Y - pt1.Y, 2), 0.5), 3);
                double bear1 = GET_Bearing_rad(pt1.X, pt1.Y, pt_prev.X, pt_prev.Y);

                if (d1 < 0.001)
                {
                    poly1.RemoveVertexAt(i);
                }
                else if (Math.Round(bear1, 3) == Math.Round(bear_prev, 3))
                {
                    poly1.RemoveVertexAt(i + 1);
                }
                else
                {
                    pt_prev = pt1;
                    bear_prev = bear1;
                }

            }
            return poly1;
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

        public static void add_od_table_to_object(ObjectId id1, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
        {
            try
            {
                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
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
                                        if (Valoare1 != null)
                                        {
                                            if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Character && Valoare1.StrValue != "")
                                            {
                                                Valoare1.Assign(List_value[i].ToString());
                                            }

                                            if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                            {
                                                Valoare1.Assign(Convert.ToDouble(List_value[i]));
                                            }

                                            if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Integer)
                                            {
                                                Valoare1.Assign(Convert.ToInt32(List_value[i]));
                                            }

                                            if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Point)
                                            {
                                                Point3d pt1 = (Point3d)List_value[i];
                                                Valoare1.Assign(pt1);
                                            }
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
                                        else if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Real)
                                        {
                                            double Valoare = 0;
                                            if (List_value[i] != null)
                                            {
                                                Valoare = Convert.ToDouble(List_value[i]);
                                            }

                                            Val.Assign(Valoare);
                                        }
                                        else if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Integer)
                                        {
                                            int Valoare = 0;
                                            if (List_value[i] != null)
                                            {
                                                Valoare = Convert.ToInt32((List_value[i]));
                                            }

                                            Val.Assign(Valoare);
                                        }
                                        else if (List_types[i] == Autodesk.Gis.Map.Constants.DataType.Point)
                                        {
                                            Point3d Valoare = new Point3d();
                                            if (List_value[i] != null)
                                            {
                                                Valoare = (Point3d)((List_value[i]));
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
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
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

            string String_minus = "";

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

            string String2, String3;
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

        public static void Transfer_datatable_to_existing_excel_spreadsheet(Worksheet W1, System.Data.DataTable dt1, string col_start, string row_start)
        {
            if (dt1 != null && W1 != null)
            {
                if (dt1.Rows.Count > 0)
                {

                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[col_start + row_start.ToString(), W1.Cells[maxRows + 1, maxCols]];
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


        public static void zoom_to_object(ObjectId ObjId)
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
                            Entity Ent1 = Trans1.GetObject(ObjId, OpenMode.ForRead) as Entity;
                            if (Ent1 != null)
                            {

                                Point3d minx = Ent1.GeometricExtents.MinPoint;
                                Point3d maxx = Ent1.GeometricExtents.MaxPoint;

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

                                            view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);

                                            view.Zoom(0.95);//<--optional 
                                            GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                        }
                                    }
                                    Trans1.Commit();
                                }

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


        public static void zoom_to_Point(Point3d pt, double factor1)
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


        public static void keep_the_view_at_pickpoint()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
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
                            //view.ZoomExtents(point_left, point_right);
                            //view.Zoom(0.95);//<--optional 
                            GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        public static Point3d get_view_center()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
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
                            return view.Position;
                        }
                    }
                    Trans1.Commit();
                }
            }
            return new Point3d();
        }

        public static ObjectId GetObjectId(Database db, string handle)
        {
            try
            {
                return db.GetObjectId(false, new Handle(Convert.ToInt64(handle)), 0);
            }
            catch (System.Exception EX)
            {

                return ObjectId.Null;
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

        static public string Get_deflection_angle_dms(double x1, double y1, double x2, double y2, double x3, double y3)
        {


            double a1 = x2 - x1;
            double b1 = y2 - y1;
            double a2 = x3 - x2;
            double b2 = y3 - y2;
            double Defl_DD = 180 * Math.Acos((a1 * a2 + b1 * b2) / (Math.Pow(a1 * a1 + b1 * b1, 0.5) * Math.Pow(a2 * a2 + b2 * b2, 0.5))) / Math.PI;

            Vector3d vector1 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));
            Vector3d vector2 = new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x3, y3, 0));
            Defl_DD = (vector2.GetAngleTo(vector1)) * 180 / Math.PI;


            double Bearing1 = 180 * Functions.GET_Bearing_rad(x1, y1, x2, y2) / Math.PI;
            double Bearing2 = 180 * Functions.GET_Bearing_rad(x2, y2, x3, y3) / Math.PI;

            String Suffix1 = " ";


            if (Bearing1 < 180)
            {

                if (Bearing2 < Bearing1 + 180 && Bearing2 > Bearing1)
                {
                    Suffix1 = " LT";
                }
                else
                {
                    Suffix1 = " RT";
                }
            }
            else
            {
                if (Bearing2 < Bearing1 && Bearing2 > Bearing1 - 180)
                {
                    Suffix1 = " RT";
                }
                else
                {
                    Suffix1 = " LT";
                }
            }

            return Get_DMS(Defl_DD, 0) + Suffix1;



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
            catch (System.Exception ex)
            {
                return "";
            }
        }


        static public System.Data.DataTable build_dt_from_excel(System.Data.DataTable dt1, Microsoft.Office.Interop.Excel.Worksheet W1, int start_row, int end_row, List<string> list_col, List<string> list_colxl)
        {


            if (W1 == null) return dt1;
            if (end_row - start_row < 0) return dt1;
            if (list_col.Count == 0) return dt1;
            if (list_col.Count != list_colxl.Count) return dt1;

            object[,] values_0 = new object[end_row - start_row + 1, 1];
            if (list_colxl[0] != "")
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[list_colxl[0] + start_row.ToString() + ":" + list_colxl[0] + end_row.ToString()];
                values_0 = range1.Value2;
                for (int k = 1; k <= values_0.Length; ++k)
                {
                    object Valoare_k = values_0[k, 1];
                    if (Valoare_k != null && dt1.Columns[list_col[0]].DataType == typeof(double))
                    {
                        Valoare_k = Valoare_k.ToString().Replace("+", "");
                        if (IsNumeric(Valoare_k.ToString()) == true)
                        {
                            Valoare_k = Convert.ToDouble(Valoare_k);
                        }
                        else
                        {
                            Valoare_k = DBNull.Value;
                        }
                    }

                    if (Valoare_k != null && dt1.Columns[list_col[0]].DataType == typeof(bool))
                    {
                        Valoare_k = Valoare_k.ToString().Replace(" ", "");
                        if (Valoare_k.ToString().ToLower() == "yes" || Valoare_k.ToString().ToLower() == "y" || Valoare_k.ToString().ToLower() == "true")
                        {

                        }
                        else
                        {
                            Valoare_k = DBNull.Value;
                        }
                    }

                    if (Valoare_k == null) Valoare_k = DBNull.Value;

                    if (Valoare_k != DBNull.Value)
                    {
                        dt1.Rows.Add();
                        if (dt1.Columns[list_col[0]].DataType == typeof(bool))
                        {
                            dt1.Rows[dt1.Rows.Count - 1][list_col[0]] = true;
                        }
                        else
                        {
                            dt1.Rows[dt1.Rows.Count - 1][list_col[0]] = Valoare_k;
                        }

                        if (dt1.Columns.Contains("Excel") == true)
                        {
                            if (dt1.Rows[dt1.Rows.Count - 1]["Excel"] == DBNull.Value)
                            {
                                dt1.Rows[dt1.Rows.Count - 1]["Excel"] = k;
                            }
                        }
                    }
                    else
                    {
                        k = values_0.Length + 1;
                    }
                }
            }


            if (list_col.Count > 1)
            {
                for (int i = 1; i < list_col.Count; ++i)
                {
                    object[,] values_i = new object[end_row - start_row + 1, 1];
                    if (list_colxl[i] != "")
                    {
                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[list_colxl[i] + start_row.ToString() + ":" + list_colxl[i] + end_row.ToString()];
                        values_i = range1.Value2;
                        for (int k = 1; k <= values_i.Length; ++k)
                        {
                            object Valoare_k = values_i[k, 1];
                            if (Valoare_k != null && dt1.Columns[list_col[i]].DataType == typeof(double))
                            {
                                Valoare_k = Valoare_k.ToString().Replace("+", "");
                                if (IsNumeric(Valoare_k.ToString()) == true)
                                {
                                    Valoare_k = Convert.ToDouble(Valoare_k);
                                }
                                else
                                {
                                    Valoare_k = DBNull.Value;
                                }
                            }


                            if (Valoare_k != null && dt1.Columns[list_col[i]].DataType == typeof(bool))
                            {
                                Valoare_k = Valoare_k.ToString().Replace(" ", "");
                                if (Valoare_k.ToString().ToLower() == "yes" || Valoare_k.ToString().ToLower() == "y" || Valoare_k.ToString().ToLower() == "true")
                                {
                                    Valoare_k = true;
                                }
                                else
                                {
                                    Valoare_k = DBNull.Value;
                                }
                            }

                            if (Valoare_k == null) Valoare_k = DBNull.Value;

                            if (k - 1 < dt1.Rows.Count)
                            {
                                dt1.Rows[k - 1][list_col[i]] = Valoare_k;
                            }

                        }
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


        static public double Get_distance1_block(BlockReference BR)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    if (prop.PropertyName == "Distance1" && prop.UnitsType == DynamicBlockReferencePropertyUnitsType.Distance)
                    {
                        return Convert.ToDouble(prop.Value);

                    }
                }
            }
            return 0;
        }


        static public void Incarca_existing_Blocks_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
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

        static public void Incarca_existing_Blocks_with_attributes_to_combobox(System.Windows.Forms.ComboBox Combo_blockname)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                Combo_blockname.Items.Clear();
                Combo_blockname.Items.Add("");
                Combo_blockname.Text = "";
                foreach (ObjectId Block_id in BlockTable_data1)
                {
                    BlockTableRecord Block1 = (BlockTableRecord)Trans1.GetObject(Block_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                    if (Block1.HasAttributeDefinitions == true)
                    {
                        if (Block1.Name.Contains("*") == false && Block1.Name.Contains("|") == false && Block1.Name.Contains("$") == false)
                        {
                            Combo_blockname.Items.Add(Block1.Name);
                        }
                    }
                }
                Trans1.Dispose();
            }
        }

        static public void Incarca_existing_Atributes_to_combobox(string BlockName, ComboBox Combo_atributes)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.BlockTable Block_table = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.BlockTable;
                Combo_atributes.Items.Clear();
                Combo_atributes.Items.Add("");
                if (BlockName != "" && Block_table != null)
                {
                    if (Block_table.Has(BlockName) == true)
                    {
                        BlockTableRecord BTrecordBlock = Trans1.GetObject(Block_table[BlockName], OpenMode.ForRead) as BlockTableRecord;
                        if (BTrecordBlock != null)
                        {
                            foreach (ObjectId Id1 in BTrecordBlock)
                            {
                                Entity ent = Trans1.GetObject(Id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                                if (ent != null)
                                {
                                    AttributeDefinition attDefinition1 = ent as AttributeDefinition;
                                    if (attDefinition1 != null)
                                    {
                                        Combo_atributes.Items.Add(attDefinition1.Tag);
                                    }
                                }
                            }
                        }
                    }
                }
                Trans1.Dispose();
            }
        }

        static public double Get_Param_Value_block(BlockReference BR, string param_name)
        {
            using (DynamicBlockReferencePropertyCollection pc = BR.DynamicBlockReferencePropertyCollection)
            {
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    if (prop.PropertyName == param_name && prop.UnitsType == DynamicBlockReferencePropertyUnitsType.Distance)
                    {
                        return Convert.ToDouble(prop.Value);

                    }
                }
            }
            return 0;
        }

    }
}
