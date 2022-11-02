using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace weld_sheet
{
    class Functions
    {

        public static Microsoft.Office.Interop.Excel.Application Excel1 = null;

        #region Global Variables
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

        #endregion
        public static string get_last_column_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int row_1)
        {
            string last_col = "";
            for (int i = 1; i < 100; ++i)
            {
                string contents = Convert.ToString(W1.Range[get_excel_column_letter(i) + row_1.ToString()].Value2);

                if (contents == "" || contents == null)
                {
                    if (i > 1)
                    {
                        last_col = get_excel_column_letter(i - 1);
                    }
                    else
                    {
                        last_col = get_excel_column_letter(i);
                    }
                    i = 100;
                }
            }

            return last_col;
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

        public static int get_last_row_excel(Microsoft.Office.Interop.Excel.Worksheet W1, string column1, int start)
        {
            int last_row = 0;
            for (int i = start; i < 30000; ++i)
            {

                string contents = Convert.ToString(W1.Range[column1 + i.ToString()].Value2);
                if (contents == "" || contents == null)
                {
                    last_row = i - 1;
                    i = 30000;
                }

            }

            return last_row;
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

        public static Microsoft.Office.Interop.Excel.Application Excel_Open()
        {
            try
            {
                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Exception ex)
            {
                Excel1 = new Microsoft.Office.Interop.Excel.Application();

            }

            Excel1.Visible = true;
            //xlApp.ActiveWindow.DisplayGridlines = false;
            return Excel1;
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


    }
}
