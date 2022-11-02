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

        public static Microsoft.Office.Interop.Excel.Application Excel1 = null;

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

        public static System.Data.DataTable Populate_data_table_with_excel_range(string table_name, System.Data.DataTable dt1, int row_start, int row_end, int col_start, int col_end, Range header_range, Range range1, string Error_Message)
        {
            int nocol1 = col_end - col_start + 1;
            object[,] array_col_names = new object[1, nocol1];
            array_col_names = header_range.Value2;

            for (int i = 1; i <= nocol1; ++i)
            {
                if (array_col_names[1, i] != null)
                {
                    if (dt1.Columns.Contains(array_col_names[1, i].ToString()) == false)
                    {
                        string cell_value = array_col_names[1, i].ToString();
                        dt1.Columns.Add(cell_value.ToString().ToUpper(), typeof(string));
                    }
                    else
                    {
                        MessageBox.Show(table_name + " Has 2 Headers With the Same Name." + "\r\n" + array_col_names[1, i].ToString(), Error_Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show(table_name + " Header is Missing Values." + "\r\n Column " + Functions.get_excel_column_letter(i), Error_Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }

            int nocol = col_end - col_start + 1;
            int norow = row_end - row_start + 1;

            for (int i = 1; i <= norow; ++i)
            {
                dt1.Rows.Add();
            }
            object[,] array_values = new object[norow, nocol];

            array_values = range1.Value2;

            for (int i = 0; i < dt1.Rows.Count; ++i)
            {
                for (int j = 0; j < dt1.Columns.Count; ++j)
                {
                    object Valoare1 = array_values[i + 1, j + 1];
                    if (Valoare1 == null) Valoare1 = DBNull.Value;
                    if (Valoare1 != DBNull.Value && Valoare1 != null)
                    {
                        Valoare1 = Convert.ToString(Valoare1).Replace("-2146826246", "");
                        if (Convert.ToString(Valoare1) == "") Valoare1 = DBNull.Value;
                    }
                    dt1.Rows[i][j] = Valoare1;
                }
            }
            dt1.Columns.Add("reporti", typeof(int));
            return dt1;
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
