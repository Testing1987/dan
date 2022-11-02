using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public class Functions
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

        public static bool is_bruno_coelho()
        {
            if (Environment.UserName.ToUpper() == "COE35585")
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


        [Autodesk.AutoCAD.Runtime.CommandMethod("hh2233")]
        public void ShowForm1()
        {
            if (is_hector_morales() == true)
            {
                wksp_tool.show_export_tab = true;
                if (isSECURE() == true)
                {
                    foreach (Form Forma1 in System.Windows.Forms.Application.OpenForms)
                    {
                        if (Forma1 is wksp_tool)
                        {
                            Forma1.Focus();
                            Forma1.WindowState = FormWindowState.Normal;
                            Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                              (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                            return;
                        }
                    }
                    try
                    {
                        wksp_tool forma2 = new wksp_tool();
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                        forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                             (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                        wksp_tool.show_export_tab = false;
                    }
                    catch (System.Exception EX)
                    {
                        MessageBox.Show(EX.Message);
                    }
                }
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
        static public double GET_Bearing_rad(double x1, double y1, double x2, double y2)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d CurentUCSmatrix = Editor1.CurrentUserCoordinateSystem;
            CoordinateSystem3d CurentUCS = CurentUCSmatrix.CoordinateSystem3d;
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            return new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0)).AngleOnPlane(Planul_curent);
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
        public static void Create_header_workspace_library(Worksheet W1, string Client, string Version, System.Data.DataTable dt1)
        {
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B6"];
            Object[,] valuesH = new object[6, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "VERSION";
            valuesH[1, 1] = Version;
            valuesH[2, 0] = "DATE CREATED";
            valuesH[2, 1] = DateTime.Now.ToString(new System.Globalization.CultureInfo("en-US"));
            valuesH[3, 0] = "USER ID";
            valuesH[3, 1] = Environment.UserName;
            valuesH[4, 0] = "Comments:";
            valuesH[5, 0] = "Do not add any columns to this table, also do not add any rows above row 8";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B4"];

            Color_border_range_inside(range1, 46);

            range1 = W1.Range["A5:I5"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 6); //yelloW

            range1 = W1.Range["A6:I6"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 3); //red

            range1 = W1.Range["C1:I4"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "WORKSPACE LIBRARY";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Color_border_range_outside(range1, 0);


            W1.Range["A5:A6"].Font.Bold = true;

            range1 = W1.Range["A7:I7"];
            Object[,] values_dt = new object[1, dt1.Columns.Count];
            if (dt1 != null && dt1.Columns.Count > 0)
            {
                for (int i = 0; i < dt1.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt1.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }

            W1.Range["A:A"].ColumnWidth = 14;
            W1.Range["B:B"].ColumnWidth = 50;
            W1.Range["C:E"].ColumnWidth = 13;
            W1.Range["F:F"].ColumnWidth = 3.14;
            W1.Range["G:I"].ColumnWidth = 13;



        }

        public static void Create_header(Worksheet W1, string Client, string Version, string title1, string last_column)
        {
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B6"];
            Object[,] valuesH = new object[6, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "VERSION";
            valuesH[1, 1] = Version;
            valuesH[2, 0] = "DATE CREATED";
            valuesH[2, 1] = DateTime.Now.ToString(new System.Globalization.CultureInfo("en-US"));
            valuesH[3, 0] = "USER ID";
            valuesH[3, 1] = Environment.UserName;
            valuesH[4, 0] = "Comments:";
            valuesH[5, 0] = "Do not add any columns to this table, also do not add any rows above row 8";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B4"];

            Color_border_range_inside(range1, 46);

            range1 = W1.Range["A5:" + last_column + "5"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 6); //yelloW

            range1 = W1.Range["A6:" + last_column + "6"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 3); //red

            range1 = W1.Range["C1:" + last_column + "4"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = title1;
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Color_border_range_outside(range1, 0);


            W1.Range["A5:A6"].Font.Bold = true;

            range1 = W1.Range["A7:" + last_column + "7"];
            Color_border_range_inside(range1, 41); //blue
            range1.Font.ColorIndex = 2;
            range1.Font.Size = 11;
            range1.Font.Bold = true;
        }

        public static void Create_header_layers(Worksheet W1, string Client, string Version, System.Data.DataTable dt1)
        {
            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B6"];
            Object[,] valuesH = new object[6, 2];
            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "VERSION";
            valuesH[1, 1] = Version;
            valuesH[2, 0] = "DATE CREATED";
            valuesH[2, 1] = DateTime.Now.ToString(new System.Globalization.CultureInfo("en-US"));
            valuesH[3, 0] = "USER ID";
            valuesH[3, 1] = Environment.UserName;
            valuesH[4, 0] = "Comments:";
            valuesH[5, 0] = "Do not add any columns to this table, also do not add any rows above row 8";
            range1.Value2 = valuesH;
            range1 = W1.Range["A1:B4"];

            Color_border_range_inside(range1, 46);

            range1 = W1.Range["A5:D5"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 6); //yelloW

            range1 = W1.Range["A6:D6"];
            range1.Merge();
            range1.MergeCells = true;
            Color_border_range_outside(range1, 3); //red

            range1 = W1.Range["C1:D4"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "LAYERS";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Color_border_range_outside(range1, 0);


            W1.Range["A5:A6"].Font.Bold = true;

            range1 = W1.Range["A7:D7"];
            Object[,] values_dt = new object[1, dt1.Columns.Count];
            if (dt1 != null && dt1.Columns.Count > 0)
            {
                for (int i = 0; i < dt1.Columns.Count; ++i)
                {
                    values_dt[0, i] = dt1.Columns[i].ColumnName;
                }
                range1.Value2 = values_dt;
                Color_border_range_inside(range1, 41); //blue
                range1.Font.ColorIndex = 2;
                range1.Font.Size = 11;
                range1.Font.Bold = true;
            }

            W1.Range["A:A"].ColumnWidth = 14;
            W1.Range["B:B"].ColumnWidth = 50;
            W1.Range["C:C"].ColumnWidth = 35;
            W1.Range["D:D"].ColumnWidth = 14;




        }

        public static void Transfer_to_worksheet_Data_table(Worksheet W1, System.Data.DataTable dt1, int Start_row, string format_cell)
        {

            int nr_col = dt1.Columns.Count;

            W1.Range["A:" + get_excel_column_letter(nr_col)].ClearContents();

            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    int NrR = dt1.Rows.Count;
                    int NrC = dt1.Columns.Count;


                    Object[,] values = new object[NrR, NrC];
                    for (int i = 0; i < NrR; ++i)
                    {
                        for (int j = 0; j < NrC; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values[i, j] = dt1.Rows[i][j];
                            }
                        }
                    }


                    Microsoft.Office.Interop.Excel.Range range0 = W1.Range[W1.Columns[1], W1.Columns[NrC]];
                    range0.ClearContents();
                    range0.UnMerge();
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range0.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];
                    range1.Cells.NumberFormat = format_cell;
                    range1.Value2 = values;
                    Color_border_range_inside(range1, 0);
                }
            }
        }

        public static void Transfer_to_worksheet_header_of_dt(Worksheet W1, System.Data.DataTable dt1, int Start_row, string format_cell)
        {

            if (dt1 != null)
            {
                if (dt1.Columns.Count > 0)
                {
                    int NrC = dt1.Columns.Count;
                    object[,] values = new object[1, NrC];
                    for (int i = 0; i < NrC; ++i)
                    {
                        values[0, i] = dt1.Columns[i].ColumnName;
                    }
                    Range range1 = W1.Range["A" + Start_row.ToString() + ":" + get_excel_column_letter(NrC) + Start_row.ToString()];
                    range1.ClearContents();
                    range1.UnMerge();
                    range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range1.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                    range1.Value2 = values;
                }
            }
        }


        public static Worksheet get_worksheet_W1(bool add_new_w1, string tabname)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel1;
                Workbook Workbook1 = null;
                Worksheet W1 = null;
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }
                if (Excel1 == null) return null;
                Excel1.Visible = true;
                for (int j = 1; j <= Excel1.Workbooks.Count; ++j)
                {
                    Workbook1 = Excel1.Workbooks[j];
                    for (int i = 1; i <= Workbook1.Worksheets.Count; ++i)
                    {
                        if (Workbook1.Worksheets[i].name == tabname)
                        {
                            W1 = Workbook1.Worksheets[i];
                            i = Workbook1.Worksheets.Count + 1;
                            j = Excel1.Workbooks.Count + 1;
                        }
                    }
                }

                if (W1 == null && add_new_w1 == true)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                    if (Workbook1.Worksheets.Count == 0)
                    {
                        W1 = Workbook1.Worksheets.Add();
                    }
                    else
                    {
                        W1 = Workbook1.Worksheets[1];
                    }
                    W1.Name = tabname;
                }
                return W1;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                return null;
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

        public static void Transfer_datatable_to_excel_spreadsheet(Worksheet W1, System.Data.DataTable dt1, int start1, bool transfer_header)
        {
            if (dt1 != null && start1 > 0)
            {
                if (dt1.Rows.Count > 0)
                {
                    int maxRows = dt1.Rows.Count;
                    int maxCols = dt1.Columns.Count;
                    int extra1 = 0;
                    if (transfer_header == true)
                    {
                        extra1 = 1;
                    }
                    object[,] values1 = new object[maxRows + extra1, maxCols];
                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i + extra1, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }
                    if (transfer_header == true)
                    {
                        for (int j = 0; j < dt1.Columns.Count; ++j)
                        {
                            values1[0, j] = dt1.Columns[j].ColumnName;
                        }
                    }
                    Range range1 = W1.Range[W1.Cells[start1, 1], W1.Cells[maxRows + extra1 + start1 - 1, maxCols]];
                    range1.Value2 = values1;
                }
            }
        }

        public static Worksheet Get_NEW_worksheet_from_Excel()
        {
            Microsoft.Office.Interop.Excel.Application Excel1;
            Microsoft.Office.Interop.Excel.Workbook Workbook1;
            try
            {
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Exception)
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

        static public void zoom_to_object(ObjectId ObjId)
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
                            Entity Ent1 = null;

                            Ent1 = Trans1.GetObject(ObjId, OpenMode.ForRead, true) as Entity;

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

        public static void load_object_data_table_name_to_combobox(ComboBox combo_name)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    combo_name.Items.Clear();
                    combo_name.Items.Add("");
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    System.Collections.Specialized.StringCollection col_names = new System.Collections.Specialized.StringCollection();
                    col_names = Tables1.GetTableNames();

                    for (int i = 0; i < col_names.Count; ++i)
                    {
                        string nume1 = col_names[i];
                        if (combo_name.Items.Contains(nume1) == false)
                        {
                            combo_name.Items.Add(nume1);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        static public void load_object_data_fieds_to_combobox(ComboBox combo_name, ComboBox combo_field)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            combo_field.Items.Clear();
            combo_field.Items.Add("");
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    if (Tables1.IsTableDefined(combo_name.Text) == true)
                    {
                        Autodesk.Gis.Map.ObjectData.Table tabla1 = Tables1[combo_name.Text];
                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = tabla1.FieldDefinitions;
                        for (int i = 0; i < Field_defs1.Count; ++i)
                        {
                            Autodesk.Gis.Map.ObjectData.FieldDefinition fielddef1 = Field_defs1[i];
                            if (combo_field.Items.Contains(fielddef1.Name) == false)
                            {
                                combo_field.Items.Add(fielddef1.Name);
                            }

                        }
                    }
                    else
                    {
                        combo_field.Items.Clear();
                    }
                    Trans1.Commit();
                }
            }
        }



        public static Autodesk.Gis.Map.ObjectData.Table Get_object_data_table(string Nume_table, string Description_table, List<string> List_Names, List<string> List_descriptions, List<Autodesk.Gis.Map.Constants.DataType> List_types, bool define_new_dt)
        {
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                return Tables1[Nume_table];
            }

            if (Tables1.IsTableDefined(Nume_table) == false && define_new_dt == true)
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

        public static void Populate_object_data_table_from_objectid(ObjectId id1, string Nume_table, List<object> List_value, List<Autodesk.Gis.Map.Constants.DataType> List_types)
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

        public static List<Autodesk.Gis.Map.Constants.DataType> get_object_data_table_data_types(Autodesk.Gis.Map.ObjectData.Tables Tables1, string Nume_table)
        {
            List<Autodesk.Gis.Map.Constants.DataType> List_types = null;

            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                List_types = new List<Autodesk.Gis.Map.Constants.DataType>();
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {

                    using (Autodesk.Gis.Map.ObjectData.FieldDefinitions field_defs = Tabla1.FieldDefinitions)
                    {
                        if (field_defs.Count > 0)
                        {
                            for (int i = 0; i < field_defs.Count; ++i)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinition field1 = field_defs[i];
                                List_types.Add(field1.Type);
                            }
                        }
                    }
                }
            }
            return List_types;
        }

        public static List<string> get_object_data_table_field_names(Autodesk.Gis.Map.ObjectData.Tables Tables1, string Nume_table)
        {
            List<string> List_names = null;

            if (Tables1.IsTableDefined(Nume_table) == true)
            {
                List_names = new List<string>();
                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Nume_table])
                {

                    using (Autodesk.Gis.Map.ObjectData.FieldDefinitions field_defs = Tabla1.FieldDefinitions)
                    {
                        if (field_defs.Count > 0)
                        {
                            for (int i = 0; i < field_defs.Count; ++i)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinition field1 = field_defs[i];
                                List_names.Add(field1.Name);
                            }
                        }
                    }
                }
            }
            return List_names;
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
        public static string Angle_left_right(Polyline Poly2D, Point3d Punct1)
        {
            String LT_RT = "";
            Point3d Point_on_poly = Poly2D.GetClosestPointTo(Punct1, Autodesk.AutoCAD.Geometry.Vector3d.ZAxis, true);
            Autodesk.AutoCAD.Geometry.Vector3d vector2 = Point_on_poly.GetVectorTo(Punct1);
            double Param1 = Poly2D.GetParameterAtPoint(Point_on_poly);
            Autodesk.AutoCAD.Geometry.Vector3d vector1;
            if (Param1 > 0)
            {
                if (Param1 == Poly2D.NumberOfVertices - 1)
                {
                    vector1 = Poly2D.GetPointAtParameter(Param1 - 1).GetVectorTo(Poly2D.GetPointAtParameter(Param1));
                }
                else
                {
                    vector1 = Poly2D.GetPointAtParameter(Math.Floor(Param1)).GetVectorTo(Poly2D.GetPointAtParameter(Math.Ceiling(Param1)));
                }
            }
            else
            {
                vector1 = Poly2D.GetPointAtParameter(0).GetVectorTo(Poly2D.GetPointAtParameter(1));
            }
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Autodesk.AutoCAD.Geometry.Vector3d.ZAxis);
            double Bearing1 = (vector1.AngleOnPlane(Planul_curent)) * 180 / Math.PI;
            double Bearing2 = (vector2.AngleOnPlane(Planul_curent)) * 180 / Math.PI;
            double angle1 = (vector2.GetAngleTo(vector1)) * 180 / Math.PI;
            if (Bearing1 < 180)
            {
                if (Bearing2 < Bearing1 + 180 && Bearing2 > Bearing1)
                {
                    LT_RT = "LT.";
                }
                else
                {
                    LT_RT = "RT.";
                }
            }
            else
            {
                if (Bearing2 < Bearing1 & Bearing2 > Bearing1 - 180)
                {
                    LT_RT = "RT.";
                }
                else
                {
                    LT_RT = "LT.";
                }
            }
            return LT_RT;
        }
        public static bool IsRightDirection(Curve pCurv, Point3d p)
        {
            Point3d pDir = (Point3d)(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("VIEWDIR"));
            Vector3d vDir = pDir.GetAsVector();

            Vector3d vNormal = Vector3d.ZAxis;
            if (pCurv.IsPlanar)
            {
                Plane plane = pCurv.GetPlane();
                vNormal = plane.Normal;
                p = p.Project(plane, vDir);
            }
            Point3d pNear = pCurv.GetClosestPointTo(p, true);
            Vector3d vSide = p - pNear;
            Vector3d vDeriv = pCurv.GetFirstDerivative(pNear);
            if (vNormal.CrossProduct(vDeriv).DotProduct(vSide) < 0.0)
                return true;
            else
                return false;
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


        public static double get_2d_distance(Point3d p1, Point3d p2)
        {
            return Math.Pow(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2), 0.5);
        }

    }
}
