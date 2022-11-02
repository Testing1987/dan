using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
//using Autodesk.AutoCAD.Runtime;
//using Autodesk.AutoCAD.Geometry;
//using Autodesk.AutoCAD.ApplicationServices;
//using Autodesk.AutoCAD.DatabaseServices;

namespace Alignment_mdi
{
    public partial class SGEN_Sheet_Index : Form
    {



        string Col_MMid = "MMID";
        string Col_handle = "AcadHandle";
        string Col_dwg_name = "DwgNo";
        string Col_M1 = "StaBeg";
        string Col_M2 = "StaEnd";
        string Col_dispM1 = "Disp_StaBeg";
        string Col_dispM2 = "Disp_StaEnd";
        string Col_length = "Length";

        string Col_Width = "Width";
        string Col_Height = "Height";
        string Col_X1 = "X_Beg";
        string Col_Y1 = "Y_Beg";
        string Col_X2 = "X_End";
        string Col_Y2 = "Y_End";

        string col_scale = "Scale";
        string col_scaleName = "ScaleName";
        public scales_form forma3 = null;

        private ContextMenuStrip ContextMenuStrip_xl;


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_place_rectangles);
            lista_butoane.Add(button_scan);
            lista_butoane.Add(button_recover_matchlines);
            lista_butoane.Add(button_save_to_excel);
            lista_butoane.Add(button_load_segment_sheet_index);
            lista_butoane.Add(button_delete_sheet_index);
            lista_butoane.Add(button_open_sheet_index_xl);
            lista_butoane.Add(dataGridView_sheet_index);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_place_rectangles);
            lista_butoane.Add(button_scan);
            lista_butoane.Add(button_recover_matchlines);
            lista_butoane.Add(button_save_to_excel);
            lista_butoane.Add(button_load_segment_sheet_index);
            lista_butoane.Add(button_delete_sheet_index);
            lista_butoane.Add(button_open_sheet_index_xl);
            lista_butoane.Add(dataGridView_sheet_index);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public SGEN_Sheet_Index()
        {
            InitializeComponent();

            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Delete Row" };
            toolStripMenuItem1.Click += delete_row_Click;

            ContextMenuStrip_xl = new ContextMenuStrip();
            ContextMenuStrip_xl.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1 });

        }

        private void delete_row_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_sheet_index.RowCount > 0)
                {
                    int index_grid = dataGridView_sheet_index.CurrentCell.RowIndex;
                    if (index_grid == -1)
                    {
                        return;
                    }

                    string dwg1_name = "";

                    if (dataGridView_sheet_index.Rows[index_grid].Cells[_SGEN_mainform.Col_dwg_name].Value != DBNull.Value)
                    {
                        dwg1_name = Convert.ToString(dataGridView_sheet_index.Rows[index_grid].Cells[_SGEN_mainform.Col_dwg_name].Value);
                    }

                    if (dwg1_name != "")
                    {
                        for (int j = _SGEN_mainform.dt_sheet_index.Rows.Count - 1; j >= 0; --j)
                        {
                            if (_SGEN_mainform.dt_sheet_index.Rows[j][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string dwg2_name = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[j][_SGEN_mainform.Col_dwg_name]);

                                if (dwg2_name.ToLower() == dwg1_name.ToLower())
                                {
                                    _SGEN_mainform.dt_sheet_index.Rows[j].Delete();
                                    label_not_saved.Visible = true;
                                    j = -1;
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

        private void dataGridView_sheet_index_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                ContextMenuStrip_xl.Show(Cursor.Position);
                ContextMenuStrip_xl.Visible = true;
            }
            else
            {
                ContextMenuStrip_xl.Visible = false;
            }
        }


        private double get_distance(Point3d pt1, Point3d pt2)
        {
            double x1 = pt1.X;
            double y1 = pt1.Y;
            double x2 = pt2.X;
            double y2 = pt2.Y;

            return Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);

        }

        public System.Data.DataTable Build_Data_table_sheet_index_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row, bool report_error = true)
        {

            System.Data.DataTable Data_table_sheet_index = Creaza_sheet_index_datatable_structure();
            string Col1 = "C";

            Range range2 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            bool is_data = false;
            for (int i = 1; i <= values2.Length; ++i)
            {
                object Valoare2 = values2[i, 1];
                if (Valoare2 != null)
                {
                    Data_table_sheet_index.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values2.Length + 1;
                }
            }

            if (is_data == false)
            {
                if (report_error == true) MessageBox.Show("no drawing numbers defined on column C of the sheet index file");
                return Data_table_sheet_index;
            }

            int NrR = Data_table_sheet_index.Rows.Count;
            int NrC = Data_table_sheet_index.Columns.Count;

            if (is_data == true)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];
                object[,] values = new object[NrR - 1, NrC - 1];
                values = range1.Value2;
                for (int i = 0; i < Data_table_sheet_index.Rows.Count; ++i)
                {
                    for (int j = 0; j < Data_table_sheet_index.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        Data_table_sheet_index.Rows[i][j] = Valoare;
                    }
                }
            }
            return Data_table_sheet_index;
        }

        public System.Data.DataTable Creaza_sheet_index_datatable_structure()
        {


            System.Type type_string = typeof(string);
            System.Type type_double = typeof(double);

            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();

            Lista1.Add(Col_MMid);
            Lista1.Add(Col_handle);
            Lista1.Add(Col_dwg_name);
            Lista1.Add(Col_M1);
            Lista1.Add(Col_M2);
            Lista1.Add(Col_dispM1);
            Lista1.Add(Col_dispM2);
            Lista1.Add(Col_length);
            Lista1.Add(_SGEN_mainform.Col_x);
            Lista1.Add(_SGEN_mainform.Col_y);
            Lista1.Add(_SGEN_mainform.Col_rot);
            Lista1.Add(Col_Width);
            Lista1.Add(Col_Height);
            Lista1.Add(Col_X1);
            Lista1.Add(Col_Y1);
            Lista1.Add(Col_X2);
            Lista1.Add(Col_Y2);
            Lista1.Add(col_scale);
            Lista1.Add(col_scaleName);

            Lista2.Add(type_string);
            Lista2.Add(type_string);
            Lista2.Add(type_string);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_double);
            Lista2.Add(type_string);

            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
        }


        public void set_dataGridView_sheet_index()
        {
            dataGridView_sheet_index.DataSource = _SGEN_mainform.dt_sheet_index;
            dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_sheet_index.EnableHeadersVisualStyles = false;
        }


        private void button_open_sheet_index_xl_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                {
                    string ProjF = _SGEN_mainform.project_main_folder;
                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }


                    string segment1 = _SGEN_mainform.tpage_settings.get_combobox_segment_name_value();

                    if (segment1 != "")
                    {
                        ProjF = ProjF + segment1 + "\\";
                    }

                    string fisier_sheet_index = ProjF + _SGEN_mainform.sheet_index_excel_name;
                    if (System.IO.File.Exists(fisier_sheet_index) == false)
                    {
                        set_enable_true();
                        MessageBox.Show("the block sheet index data file does not exist");
                        return;
                    }
                    Microsoft.Office.Interop.Excel.Application Excel1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();

                    }
                    Excel1.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_sheet_index);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();
        }

        private void button_delete_sheet_index_Click(object sender, EventArgs e)
        {
            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }

            if (MessageBox.Show("WARNING!!!\r\nThis will remove all sheet indexes from your drawing and remove all data from the Sheet Index Data Table.\r\nDo you want to continue?", "AGEN", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;

            }


            set_enable_false();
            try
            {

                _SGEN_mainform.dt_sheet_index = Creaza_sheet_index_datatable_structure();
                dataGridView_sheet_index.DataSource = _SGEN_mainform.dt_sheet_index;
                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_sheet_index.EnableHeadersVisualStyles = false;

                if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                {
                    string ProjF = _SGEN_mainform.project_main_folder;
                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }
                    string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;

                    Functions.create_backup(fisier_si);
                    Populate_sheet_index_file(fisier_si);

                    Erase_matchlines_templates();


                    label_not_saved.Visible = false;

                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            set_enable_true();
        }

        public void Populate_sheet_index_file(string File1)
        {
            try
            {
                bool is_file_open = false;
                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(File1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }

                else
                {
                    foreach (Workbook wk1 in Excel1.Workbooks)
                    {
                        if (wk1.FullName.ToLower() == File1.ToLower())
                        {
                            Workbook1 = wk1;
                            is_file_open = true;
                        }
                    }


                    if (Workbook1 == null) Workbook1 = Excel1.Workbooks.Open(File1);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    string segment1 = "";

                    if (_SGEN_mainform.dt_sheet_index != null && _SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; i++)
                        {
                            if (_SGEN_mainform.dt_sheet_index.Rows[i][col_scale] != DBNull.Value)
                            {
                                if (_SGEN_mainform.dt_sheet_index.Rows[i][col_scaleName] != DBNull.Value)
                                {
                                    string scalename2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[i][col_scaleName]);
                                    if (scalename2.Contains(":") == true)
                                    {
                                        string[] text_array1 = scalename2.Split(Convert.ToChar(":"));
                                        if (text_array1.Length == 2)
                                        {
                                            string numarator = text_array1[0];
                                            string numitor = text_array1[1];
                                            if (Functions.IsNumeric(numarator) == true && Functions.IsNumeric(numitor) == true)
                                            {
                                                double nr2 = Convert.ToDouble(numarator);
                                                double nrr2 = Convert.ToDouble(numitor);
                                                if (nr2 > 0 && nrr2 > 0)
                                                {
                                                    _SGEN_mainform.dt_sheet_index.Rows[i][col_scaleName] = "'" + scalename2;
                                                }
                                            }
                                        }
                                    }
                                }
                            }




                        }
                    }


                    Functions.Transfer_to_worksheet_Data_table(W1, _SGEN_mainform.dt_sheet_index, _SGEN_mainform.Start_row_Sheet_index, "General");
                    Create_header_sheet_index_file(W1, _SGEN_mainform.tpage_settings.Get_client_name(), _SGEN_mainform.tpage_settings.Get_project_name(), segment1);

                    W1.Range["A:A"].ColumnWidth = 15;
                    W1.Range["B:C"].ColumnWidth = 20;
                    W1.Range["D:H"].ColumnWidth = 2;
                    W1.Range["I:M"].ColumnWidth = 15;
                    W1.Range["N:Q"].ColumnWidth = 2;
                    W1.Range["R:S"].ColumnWidth = 10;
                    W1.Name = "SI_" + System.DateTime.Now.Year.ToString() + "_" + System.DateTime.Now.Month.ToString() + "_" + System.DateTime.Now.Day.ToString();
                    if (System.IO.File.Exists(File1) == false)
                    {
                        Workbook1.SaveAs(File1);
                    }
                    else
                    {
                        Workbook1.Save();
                    }
                    if (is_file_open == false) Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Erase_matchlines_templates()
        {
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
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Functions.delete_entities_with_OD(_SGEN_mainform.Layer_name_ML_rectangle, _SGEN_mainform.od_table_sheet_index);

                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

        }
        public static void Create_header_sheet_index_file(Worksheet W1, string Client, string Project, string Segment)
        {


            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:S10"];


            Object[,] valuesH = new object[10, 19];

            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";
            valuesH[3, 1] = "";
            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;
            valuesH[6, 0] = "If this data is manually edited, the sheet indexes in the basefile must be re-drawn.";
            valuesH[7, 0] = "Do not add any columns to this table, also do not add any rows above row 12";
            valuesH[8, 0] = "Only green columns can be edited (user):";
            valuesH[9, 0] = "n/a";
            valuesH[9, 1] = "Program";
            valuesH[9, 2] = "User";
            valuesH[9, 3] = "User";
            valuesH[9, 4] = "User";
            valuesH[9, 5] = "User";
            valuesH[9, 6] = "User";
            valuesH[9, 7] = "Program";
            valuesH[9, 8] = "Program";
            valuesH[9, 9] = "Program";
            valuesH[9, 10] = "Program";
            valuesH[9, 11] = "Program";
            valuesH[9, 12] = "User";
            valuesH[9, 13] = "User";
            valuesH[9, 14] = "User";
            valuesH[9, 15] = "User";
            valuesH[9, 16] = "User";
            valuesH[9, 17] = "Program";
            valuesH[9, 18] = "Program";
            range1.Value2 = valuesH;

            range1 = W1.Range["A1:B6"];
            Functions.Color_border_range_inside(range1, 46);

            range1 = W1.Range["A7:S7"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 6);

            range1 = W1.Range["A8:S8"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 3);

            range1 = W1.Range["A9:S9"];
            range1.Merge();
            range1.MergeCells = true;
            Functions.Color_border_range_outside(range1, 43);

            range1 = W1.Range["A10:S10"];
            Functions.Color_border_range_inside(range1, 43);

            W1.Range["B10:B10"].Interior.ColorIndex = 16;
            W1.Range["H10:M10"].Interior.ColorIndex = 16;

            range1 = W1.Range["C1:S6"];
            range1.Merge();
            range1.MergeCells = true;
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range1.Value2 = "SheetIndex";
            range1.Font.Name = "Arial Black";
            range1.Font.Size = 20;
            range1.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            Functions.Color_border_range_outside(range1, 0);

            range1 = W1.Range["A11:S11"];
            range1.Font.Color = 16777215;
            range1.Font.Bold = true;
            range1.Interior.ColorIndex = 41;
        }

        private void button_load_segment_sheet_index_Click(object sender, EventArgs e)
        {
            Build_sheet_index_dt_from_excel();
        }
        public void Build_sheet_index_dt_from_excel(string sheetname = "")
        {
            string ProjF = _SGEN_mainform.project_main_folder;

            string segment1 = _SGEN_mainform.tpage_settings.get_combobox_segment_name_value();

            if (segment1 == "")
            {
                if (_SGEN_mainform.no_of_segments > 0)
                {
                    if (_SGEN_mainform.dt_segments.Rows[0]["Segment Name"] != DBNull.Value)
                    {
                        segment1 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[0]["Segment Name"]);
                    }
                }
            }

            if (System.IO.Directory.Exists(ProjF) == true)
            {

                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                if (segment1 != "")
                {
                    if (System.IO.Directory.Exists(ProjF + segment1) == true)
                    {
                        ProjF = ProjF + segment1 + "\\";
                    }
                }



                bool excel_is_opened = false;
                string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;


                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == fisier_si.ToLower())
                        {
                            Workbook1 = Workbook2;

                            if (sheetname == "")
                            {
                                W1 = Workbook1.Worksheets[1];
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                                {
                                    if (W2.Name.ToLower() == sheetname.ToLower())
                                    {
                                        W1 = W2;
                                    }
                                }
                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets[1];
                                }
                            }

                            excel_is_opened = true;
                        }

                    }

                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }


                if (System.IO.File.Exists(fisier_si) == true)
                {
                    if (W1 == null)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                        Workbook1 = Excel1.Workbooks.Open(fisier_si);

                        foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                        {
                            if (W2.Name.ToLower() == sheetname.ToLower())
                            {
                                W1 = W2;
                            }
                        }
                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets[1];
                        }

                    }
                }

                try
                {

                    if (System.IO.File.Exists(fisier_si) == true)
                    {


                        _SGEN_mainform.dt_sheet_index = Build_Data_table_sheet_index_from_excel(W1, _SGEN_mainform.Start_row_Sheet_index + 1);
                        _SGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();
                        if (excel_is_opened == false)
                        {
                            Workbook1.Close();
                        }


                    }





                    if (excel_is_opened == false)
                    {
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }

                    }

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }

                if (System.IO.File.Exists(fisier_si) == false)
                {
                    _SGEN_mainform.dt_sheet_index = null;
                    _SGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();
                }


            }
            else
            {
                MessageBox.Show("the Project database folder location is not specified\r\n" + ProjF + "\r\n operation aborted");

                return;
            }
        }

        private void button_recover_matchlines_Click(object sender, EventArgs e)
        {




            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            set_enable_false();

            if (_SGEN_mainform.dt_sheet_index != null)
            {
                if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    Create_ML_object_dataTABLE();

                    try
                    {
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {


                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);




                                delete_polyline_with_OD(_SGEN_mainform.Layer_name_ML_rectangle, _SGEN_mainform.od_table_sheet_index);
                                delete_mtext_with_OD(_SGEN_mainform.Layer_name_ML_rectangle, _SGEN_mainform.od_table_sheet_index);


                                int CI = 1;
                                if (checkBox_plat_mode.Checked == true)
                                {
                                    CI = 256;
                                }
                                Functions.Creaza_layer(_SGEN_mainform.Layer_name_ML_rectangle, 4, false);
                                Polyline Rectanglep = null;

                                for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {



                                    double Cx = 0;
                                    double Cy = 0;
                                    double rotation = 0;
                                    double width1 = 0;
                                    double height1 = 0;

                                    double Cxn = 0;
                                    double Cyn = 0;
                                    double rotationn = 0;
                                    double width1n = 0;
                                    double height1n = 0;

                                    if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_x] != DBNull.Value)
                                    {
                                        Cx = (double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_x];
                                    }
                                    else
                                    {

                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle X value for sheet index in row " + (i).ToString());
                                        return;
                                    }
                                    if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_y] != DBNull.Value)
                                    {
                                        Cy = (double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_y];
                                    }
                                    else
                                    {

                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Y value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_rot] != DBNull.Value)
                                    {
                                        rotation = (double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_rot] * Math.PI / 180;
                                    }
                                    else
                                    {

                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle ROTATION value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_SGEN_mainform.dt_sheet_index.Rows[i][Col_Height] != DBNull.Value)
                                    {
                                        height1 = (double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_Height];
                                    }
                                    else
                                    {

                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Height value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (_SGEN_mainform.dt_sheet_index.Rows[i][Col_Width] != DBNull.Value)
                                    {
                                        width1 = (double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_Width];
                                    }
                                    else
                                    {

                                        set_enable_true();
                                        MessageBox.Show("no matchline rectangle Width value for sheet index in row " + (i).ToString());
                                        return;
                                    }

                                    if (i < _SGEN_mainform.dt_sheet_index.Rows.Count - 1)
                                    {


                                        if (_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_x] != DBNull.Value)
                                        {
                                            Cxn = (double)_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_x];
                                        }
                                        else
                                        {

                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle X value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }
                                        if (_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_y] != DBNull.Value)
                                        {
                                            Cyn = (double)_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_y];
                                        }
                                        else
                                        {

                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Y value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_rot] != DBNull.Value)
                                        {
                                            rotationn = (double)_SGEN_mainform.dt_sheet_index.Rows[i + 1][_SGEN_mainform.Col_rot] * Math.PI / 180;
                                        }
                                        else
                                        {

                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle ROTATION value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_SGEN_mainform.dt_sheet_index.Rows[i + 1][Col_Height] != DBNull.Value)
                                        {
                                            height1n = (double)_SGEN_mainform.dt_sheet_index.Rows[i + 1][Col_Height];
                                        }
                                        else
                                        {

                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Height value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }

                                        if (_SGEN_mainform.dt_sheet_index.Rows[i + 1][Col_Width] != DBNull.Value)
                                        {
                                            width1n = (double)_SGEN_mainform.dt_sheet_index.Rows[i + 1][Col_Width];
                                        }
                                        else
                                        {

                                            set_enable_true();
                                            MessageBox.Show("no matchline rectangle Width value for sheet index in row " + (i + 1).ToString());
                                            return;
                                        }
                                    }





                                    Polyline Rectangle1 = creaza_rectangle_from_one_point(new Point3d(Cx, Cy, 0), rotation, width1, height1, CI);
                                    Rectangle1.Layer = _SGEN_mainform.Layer_name_ML_rectangle;



                                    BTrecord.AppendEntity(Rectangle1);
                                    Trans1.AddNewlyCreatedDBObject(Rectangle1, true);

                                    _SGEN_mainform.dt_sheet_index.Rows[i][Col_handle] = Rectangle1.ObjectId.Handle.Value.ToString();

                                    if (checkBox_plat_mode.Checked == false)
                                    {
                                        CI = CI + 1;
                                        if (CI > 7) CI = 1;
                                    }




                                    Rectanglep = Rectangle1;
                                }


                                Append_ML_object_data();

                                Trans1.Commit();
                                dataGridView_sheet_index.DataSource = _SGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;


                                label_not_saved.Visible = false;

                            }



                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }

                set_enable_true();
            }

        }

        public void delete_polyline_with_OD(string layer_name, string od_table_name)
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
                        Polyline ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
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

        public void delete_mtext_with_OD(string layer_name, string od_table_name)
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
                        MText ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
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
        public Polyline creaza_rectangle_from_one_point(Point3d Point1, double Rotation_rad, double Width1, double Height1, int cid)
        {
            Polyline Poly1r = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            Poly1r.AddVertexAt(0, new Point2d(Point1.X - Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(1, new Point2d(Point1.X - Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(2, new Point2d(Point1.X + Width1 / 2, Point1.Y + Height1 / 2), 0, 0, 0);
            Poly1r.AddVertexAt(3, new Point2d(Point1.X + Width1 / 2, Point1.Y - Height1 / 2), 0, 0, 0);


            Poly1r.Closed = true;
            Poly1r.ColorIndex = cid;

            Poly1r.TransformBy(Matrix3d.Rotation(Rotation_rad, Vector3d.ZAxis, Point1));

            return Poly1r;
        }



        private void Append_ML_object_data()
        {

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

                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            delete_mtext_with_OD(_SGEN_mainform.Layer_name_ML_rectangle, _SGEN_mainform.od_table_sheet_index);




                            for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                            {

                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                string ObjID = _SGEN_mainform.dt_sheet_index.Rows[i][Col_handle].ToString();

                                ObjectId id_poly = Functions.GetObjectId(ThisDrawing.Database, ObjID);
                                if (id_poly != ObjectId.Null)
                                {
                                    Lista_val.Add(ObjID);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Polyline Poly1 = Trans1.GetObject(id_poly, OpenMode.ForWrite) as Polyline;
                                    if (Poly1 != null)
                                    {
                                        string nume_dwg = _SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name].ToString();

                                        MText Mt1_label = new MText();
                                        Mt1_label.Contents = nume_dwg;

                                        Mt1_label.TextHeight = _SGEN_mainform.Vw_height / _SGEN_mainform.Vw_scale / 10;

                                        Mt1_label.Rotation = Functions.GET_Bearing_rad(Poly1.GetPointAtParameter(1).X, Poly1.GetPointAtParameter(1).Y, Poly1.GetPointAtParameter(2).X, Poly1.GetPointAtParameter(2).Y);
                                        Mt1_label.Attachment = AttachmentPoint.BottomLeft;
                                        Mt1_label.Location = Poly1.GetPointAtParameter(1);


                                        Mt1_label.ColorIndex = Poly1.ColorIndex;
                                        Mt1_label.Layer = _SGEN_mainform.Layer_name_ML_rectangle;
                                        BTrecord.AppendEntity(Mt1_label);
                                        Trans1.AddNewlyCreatedDBObject(Mt1_label, true);


                                        string noname = "NOT ASSIGNED";
                                        double ZERO = 0;

                                        if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                                        {
                                            Lista_val.Add(_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name].ToString());
                                        }
                                        else
                                        {
                                            Lista_val.Add(noname);
                                        }
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                                        if (_SGEN_mainform.dt_sheet_index.Rows[i][Col_M1] != DBNull.Value)
                                        {
                                            Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_M1]);
                                        }
                                        else
                                        {
                                            Lista_val.Add(ZERO);
                                        }

                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);


                                        if (_SGEN_mainform.dt_sheet_index.Rows[i][Col_M2] != DBNull.Value)
                                        {
                                            Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_M2]);
                                        }
                                        else
                                        {
                                            Lista_val.Add(ZERO);
                                        }

                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_x]);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_y]);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_rot]);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_Width]);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add((double)_SGEN_mainform.dt_sheet_index.Rows[i][Col_Height]);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                        Lista_val.Add("Alignment Sheet");
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                        Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                        Lista_val.Add("");
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                        Lista_val.Add("");
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                        string segment1 = "";

                                        Lista_val.Add(segment1);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                                        Functions.Populate_object_data_table_from_handle_string(Tables1, ObjID, _SGEN_mainform.od_table_sheet_index, Lista_val, Lista_type);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, Mt1_label.ObjectId, _SGEN_mainform.od_table_sheet_index, Lista_val, Lista_type);
                                    }


                                }
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

        }

        private void button_save_to_excel_Click(object sender, EventArgs e)
        {

            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
            {
                string ProjF = _SGEN_mainform.project_main_folder;
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string segment1 = _SGEN_mainform.tpage_settings.get_combobox_segment_name_value();

                if (segment1 != "")
                {
                    ProjF = ProjF + segment1 + "\\";
                }

                string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;

                Functions.create_backup(fisier_si);
                Populate_sheet_index_file(fisier_si);
                label_not_saved.Visible = false;
            }
        }

        private void button_scan_Click(object sender, EventArgs e)
        {
            if (Functions.Get_if_workbook_is_open_in_Excel("sheet_index.xlsx") == true)
            {
                MessageBox.Show("Please close the sheet index file");
                return;
            }
            string ProjF = _SGEN_mainform.project_main_folder;
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }
            if (System.IO.Directory.Exists(ProjF) == false)
            {
                MessageBox.Show("no project loaded");
                return;
            }


            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Editor1.SetImpliedSelection(Empty_array);
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect rectangles viewport creation:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status == PromptStatus.OK)
                        {

                            #region read existing sheet index
                            string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;
                            if (System.IO.File.Exists(fisier_si) == true)
                            {
                                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                                try
                                {
                                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                                }
                                catch (System.Exception ex)
                                {
                                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                                }

                                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                                Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                                Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                                try
                                {
                                    if (Excel1 == null)
                                    {
                                        MessageBox.Show("PROBLEM WITH EXCEL!");
                                        return;
                                    }

                                    Workbook2 = Excel1.Workbooks.Open(fisier_si);
                                    W2 = Workbook2.Worksheets[1];
                                    _SGEN_mainform.dt_sheet_index = Build_Data_table_sheet_index_from_excel(W2, _SGEN_mainform.Start_row_Sheet_index + 1, false);
                                    Workbook2.Close();
                                    if (Excel1.Workbooks.Count == 0)
                                    {
                                        Excel1.Quit();
                                    }
                                    else
                                    {
                                        Excel1.Visible = true;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                                finally
                                {

                                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                                    if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                                }
                            }
                            else
                            {
                                _SGEN_mainform.dt_sheet_index = Creaza_sheet_index_datatable_structure();
                            }
                            #endregion




                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                Polyline rect1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                if (rect1 != null)
                                {
                                    if (rect1.NumberOfVertices > 5 || rect1.NumberOfVertices < 4)
                                    {
                                        MessageBox.Show("The rectangle can only have 4 or 5 vertices\r\nFix your rectangle and try again");
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        set_enable_true();

                                        return;
                                    }

                                    Point3d pt1 = new Point3d(0, 1, 2);
                                    Point3d pt2 = new Point3d(0, 1, 2);
                                    Point3d pt3 = new Point3d(0, 1, 2);
                                    Point3d pt4 = new Point3d(0, 1, 2);

                                    if (rect1.NumberOfVertices == 4)
                                    {
                                        pt1 = rect1.GetPointAtParameter(0);
                                        pt2 = rect1.GetPointAtParameter(1);
                                        pt3 = rect1.GetPointAtParameter(2);
                                        pt4 = rect1.GetPointAtParameter(3);
                                    }

                                    if (rect1.NumberOfVertices == 5)
                                    {
                                        pt1 = rect1.GetPointAtParameter(0);
                                        pt2 = rect1.GetPointAtParameter(1);
                                        pt3 = rect1.GetPointAtParameter(2);
                                        pt4 = rect1.GetPointAtParameter(3);

                                        Point3d pt5 = rect1.GetPointAtParameter(4);

                                        double dist1 = get_distance(pt1, pt2);
                                        double dist2 = get_distance(pt1, pt3);
                                        double dist3 = get_distance(pt1, pt4);
                                        double dist4 = get_distance(pt2, pt3);
                                        double dist5 = get_distance(pt2, pt4);
                                        double dist6 = get_distance(pt3, pt4);

                                        if (Math.Round(dist1, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist2, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist3, 0) == 0)
                                        {
                                            pt1 = pt5;
                                        }
                                        if (Math.Round(dist4, 0) == 0)
                                        {
                                            pt2 = pt5;
                                        }
                                        if (Math.Round(dist5, 0) == 0)
                                        {
                                            pt2 = pt5;
                                        }
                                        if (Math.Round(dist6, 0) == 0)
                                        {
                                            pt3 = pt5;
                                        }
                                    }

                                    if (pt1 != new Point3d(0, 1, 2) && pt2 != new Point3d(0, 1, 2) && pt3 != new Point3d(0, 1, 2) && pt4 != new Point3d(0, 1, 2))
                                    {
                                        double dist1 = get_distance(pt1, pt2);
                                        double dist2 = get_distance(pt2, pt3);
                                        double dist3 = get_distance(pt3, pt4);
                                        double dist4 = get_distance(pt4, pt1);
                                        if (Math.Abs(dist1 - dist3) < 5 && Math.Abs(dist2 - dist4) < 5)
                                        {
                                            // i assume 2 to 3 is the length 1 to 3 is the width
                                            double bear1 = Functions.GET_Bearing_rad(pt2.X, pt2.Y, pt3.X, pt3.Y);
                                            double xc = (pt1.X + pt3.X) / 2;
                                            double yc = (pt1.Y + pt3.Y) / 2;
                                            double l1 = dist2;
                                            double h1 = dist1;


                                            Polyline poly_start = new Polyline();
                                            poly_start.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                            poly_start.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                            poly_start.Elevation = 0;

                                            Polyline poly_end = new Polyline();
                                            poly_end.AddVertexAt(0, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                            poly_end.AddVertexAt(1, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                            poly_end.Elevation = 0;

                                            Polyline poly_top = new Polyline();
                                            poly_top.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                            poly_top.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                            poly_top.Elevation = 0;

                                            Polyline poly_bottom = new Polyline();
                                            poly_bottom.AddVertexAt(0, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                            poly_bottom.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                            poly_bottom.Elevation = 0;




                                            if (dist1 > dist2)
                                            {
                                                // i assume 1 to 2 is the length 2 to 3 is the height
                                                bear1 = Functions.GET_Bearing_rad(pt1.X, pt1.Y, pt2.X, pt2.Y);
                                                xc = (pt1.X + pt3.X) / 2;
                                                yc = (pt1.Y + pt3.Y) / 2;
                                                l1 = dist1;
                                                h1 = dist2;

                                                poly_start = new Polyline();
                                                poly_start.AddVertexAt(0, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                                poly_start.AddVertexAt(1, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                                poly_start.Elevation = 0;

                                                poly_end = new Polyline();
                                                poly_end.AddVertexAt(0, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                                poly_end.AddVertexAt(1, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_end.Elevation = 0;

                                                poly_top = new Polyline();
                                                poly_top.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
                                                poly_top.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
                                                poly_top.Elevation = 0;

                                                poly_bottom = new Polyline();
                                                poly_bottom.AddVertexAt(0, new Point2d(pt3.X, pt3.Y), 0, 0, 0);
                                                poly_bottom.AddVertexAt(1, new Point2d(pt4.X, pt4.Y), 0, 0, 0);
                                                poly_bottom.Elevation = 0;



                                            }

                                            string dwg_name = "";

                                            #region object data
                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), rect1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                            {
                                                if (Records1 != null)
                                                {
                                                    if (Records1.Count > 0)
                                                    {

                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                            for (int j = 0; j < Record1.Count; ++j)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                                string Nume_field = Field_def1.Name;
                                                                object valoare1 = Record1[j].StrValue;
                                                                if (Nume_field.ToLower() == "drawingnum")
                                                                {
                                                                    dwg_name = Convert.ToString(valoare1);
                                                                    j = Record1.Count;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            int index_si = _SGEN_mainform.dt_sheet_index.Rows.Count;

                                            bool possibly_replace_sheet_index = false;

                                            if (dwg_name != "")
                                            {
                                                for (int k = 0; k < _SGEN_mainform.dt_sheet_index.Rows.Count; ++k)
                                                {
                                                    if (_SGEN_mainform.dt_sheet_index.Rows[k]["DwgNo"] != DBNull.Value)
                                                    {
                                                        string dwg2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[k]["DwgNo"]);


                                                        if (dwg2.ToLower() == dwg_name.ToLower())
                                                        {
                                                            index_si = k;
                                                            possibly_replace_sheet_index = true;
                                                            k = _SGEN_mainform.dt_sheet_index.Rows.Count;
                                                        }
                                                    }
                                                }
                                            }





                                            if (possibly_replace_sheet_index == false)
                                            {
                                                _SGEN_mainform.dt_sheet_index.Rows.Add();
                                            }
                                            else
                                            {
                                                if (MessageBox.Show(dwg_name + " \r\nhas been found in sheet index.\r\nDo you want to replace it?", "agen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                                {
                                                    _SGEN_mainform.dt_sheet_index.Rows.Add();
                                                    index_si = _SGEN_mainform.dt_sheet_index.Rows.Count - 1;
                                                }
                                            }

                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["Rotation"] = bear1 * 180 / Math.PI;
                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["X"] = xc;
                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["Y"] = yc;
                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["DwgNo"] = dwg_name;

                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["Height"] = h1;
                                            _SGEN_mainform.dt_sheet_index.Rows[index_si]["Width"] = l1;

                                        }

                                    }

                                }
                            }

                            if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                            {
                                Functions.create_backup(fisier_si);
                                Populate_sheet_index_file(fisier_si);

                                dataGridView_sheet_index.DataSource = _SGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;

                            }

                            Trans1.Commit();

                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }




        private void button_place_rectangles_Click(object sender, EventArgs e)
        {





            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }




            if (_SGEN_mainform.Vw_height == 0 || _SGEN_mainform.Vw_width == 0)
            {
                MessageBox.Show("you do not have the dimensions for the matchline rectangles\r\nOperation aborted");


                return;
            }



            Create_ML_object_dataTABLE();


            this.MdiParent.WindowState = FormWindowState.Minimized;


            set_enable_false();


            ObjectId[] Empty_array = null;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    Functions.Creaza_layer(_SGEN_mainform.Layer_name_ML_rectangle, 4, false);



                    if (_SGEN_mainform.dt_sheet_index == null)
                    {
                        _SGEN_mainform.dt_sheet_index = Creaza_sheet_index_datatable_structure();
                    }

                    string Scale1 = _SGEN_mainform.tpage_settings.Get_combobox_viewport_scale_text();



                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _SGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                    }
                    else
                    {
                        if (Scale1.Contains(":") == true)
                        {
                            Scale1 = Scale1.Replace("1:", "");
                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _SGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                            }
                        }
                        else
                        {
                            string inch = "\u0022";

                            if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                            {
                                Scale1 = Scale1.Replace("1" + inch + "=", "");
                                Scale1 = Scale1.Substring(0, Scale1.Length - 1);
                            }

                            inch = "\u0094";

                            if (Scale1.Contains(inch + "=") == true && Scale1.Contains("'") == true)
                            {
                                Scale1 = Scale1.Replace("1" + inch + "=", "");
                                Scale1 = Scale1.Substring(0, Scale1.Length - 1);

                            }

                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _SGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                            }
                        }
                    }



                    double scale2 = _SGEN_mainform.Vw_scale;
                    string nume2 = "1:" + Convert.ToString(1 / scale2);

                    int Colorindex = 1;
                    if (checkBox_plat_mode.Checked == true) Colorindex = 256;

                    string anchor = "TL";


                    Autodesk.AutoCAD.EditorInput.PromptKeywordOptions Prompt_string = new Autodesk.AutoCAD.EditorInput.PromptKeywordOptions("");
                    Prompt_string.Message = "\nSpecify ANCHOR:";

                    if (checkBox_plat_mode.Checked == true)
                    {
                        Prompt_string.Keywords.Add("CEN");
                    }
                    Prompt_string.Keywords.Add("TL");
                    Prompt_string.Keywords.Add("TR");
                    Prompt_string.Keywords.Add("BL");
                    Prompt_string.Keywords.Add("BR");

                    if (checkBox_plat_mode.Checked == true)
                    {
                        Prompt_string.Keywords.Default = "CEN";
                    }
                    else
                    {
                        Prompt_string.Keywords.Default = "TL";
                    }

                    Prompt_string.AllowNone = true;

                    if (checkBox_plat_mode.Checked == false)
                    {
                        Autodesk.AutoCAD.EditorInput.PromptResult Rezultat_suffix = ThisDrawing.Editor.GetKeywords(Prompt_string);

                        if (Rezultat_suffix.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK || Rezultat_suffix.StringResult == "")
                        {

                            anchor = "TL";

                        }
                        else
                        {
                            anchor = Rezultat_suffix.StringResult;
                        }
                    }
                    else
                    {
                        anchor = "CEN";
                    }




                    bool run1 = true;


                    do
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            BlockTable BlockTable1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Result_point1 = null;



                            //Alignment_mdi.Jig_show_rectangle_top_left Jig_top_left = new Alignment_mdi.Jig_show_rectangle_top_left();
                            //Result_point1 = Jig_top_left.StartJig(_SGEN_mainform.Vw_scale, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, anchor);


                            Jig_show_rectangle_with_JigPromptPointOptions jig1 = new Jig_show_rectangle_with_JigPromptPointOptions();

                            if (checkBox_plat_mode.Checked == true)
                            {
                                scale2 = forma3.get_current_scale();
                            }

                            if (checkBox_plat_mode.Checked == true)
                            {
                                Result_point1 = jig1.StartJig(scale2, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, anchor, true);
                            }
                            else
                            {
                                Result_point1 = jig1.StartJig(scale2, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, anchor);
                            }

                            if (Result_point1 == null || Result_point1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel)
                            {
                                run1 = false;
                                Editor1.WriteMessage("\nCommand:");


                            }

                            if (run1 == true)
                            {

                                if (checkBox_plat_mode.Checked == true)
                                {
                                    scale2 = forma3.get_current_scale();
                                    nume2 = forma3.get_current_scale_name();
                                }



                                Point3d Point1 = Result_point1.Value;
                                Polyline Rectangle2 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

                                Rectangle2 = create_rectangle_VP(scale2, Point1, anchor, Colorindex);
                                Rectangle2.Layer = _SGEN_mainform.Layer_name_ML_rectangle;

                                BTrecord.AppendEntity(Rectangle2);
                                Trans1.AddNewlyCreatedDBObject(Rectangle2, true);

                                bool add_row = true;
                                if (checkBox_plat_mode.Checked == true)
                                {

                                    if (MessageBox.Show("Is this Ok?", "Plats", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                                    {
                                        Rectangle2.Erase();
                                        add_row = false;
                                    }

                                }

                                Trans1.Commit();




                                if (add_row == true)
                                {

                                    _SGEN_mainform.dt_sheet_index.Rows.Add();

                                    if (checkBox_plat_mode.Checked == true)
                                    {
                                        if (checkBox_pick_name_from_OD.Checked == true)
                                        {
                                            if (comboBox_od_field.Text != "")
                                            {

                                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_entity;
                                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_entity;
                                                Prompt_entity = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the entity containing object data:");
                                                Prompt_entity.SetRejectMessage("\nSelect an entity!");
                                                Prompt_entity.AllowNone = true;
                                                Prompt_entity.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Entity), false);
                                                Rezultat_entity = ThisDrawing.Editor.GetEntity(Prompt_entity);

                                                string Name_of_sheet = "XXX";

                                                if (Rezultat_entity.Status == PromptStatus.OK)
                                                {
                                                    #region object data

                                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Rezultat_entity.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                    {
                                                        if (Records1 != null)
                                                        {
                                                            if (Records1.Count > 0)
                                                            {

                                                                foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                                    for (int j = 0; j < Record1.Count; ++j)
                                                                    {
                                                                        Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                                        string Nume_field = Field_def1.Name;
                                                                        object valoare1 = Record1[j].StrValue;
                                                                        if (Nume_field == comboBox_od_field.Text)
                                                                        {
                                                                            Name_of_sheet = Convert.ToString(valoare1);
                                                                            j = Record1.Count;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_dwg_name] = Name_of_sheet;

                                                }

                                            }
                                        }
                                    }



                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_handle] = Rectangle2.ObjectId.Handle.Value.ToString();
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_x] = (Rectangle2.GetPoint3dAt(0).X + Rectangle2.GetPoint3dAt(2).X) / 2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_y] = (Rectangle2.GetPoint3dAt(0).Y + Rectangle2.GetPoint3dAt(2).Y) / 2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][_SGEN_mainform.Col_rot] = Functions.GET_Bearing_rad(Rectangle2.GetPoint3dAt(1).X, Rectangle2.GetPoint3dAt(1).Y, Rectangle2.GetPoint3dAt(2).X, Rectangle2.GetPoint3dAt(2).Y) * 180 / Math.PI;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_Width] = Rectangle2.GetPoint3dAt(1).DistanceTo(Rectangle2.GetPoint3dAt(2));
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][Col_Height] = Rectangle2.GetPoint3dAt(0).DistanceTo(Rectangle2.GetPoint3dAt(1));
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][col_scale] = scale2;
                                    _SGEN_mainform.dt_sheet_index.Rows[_SGEN_mainform.dt_sheet_index.Rows.Count - 1][col_scaleName] = nume2;
                                    if (checkBox_plat_mode.Checked == false)
                                    {
                                        Colorindex = Colorindex + 1;
                                        if (Colorindex > 7) Colorindex = 1;
                                    }
                                }



                            }
                        }
                    } while (run1 == true);




                    if (_SGEN_mainform.dt_sheet_index != null)
                    {
                        if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {
                            if (checkBox_pick_name_from_OD.Checked == false)
                            {
                                Populate_data_table_matchline_file_names();
                            }

                            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                            {
                                string ProjF = _SGEN_mainform.project_main_folder;
                                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                {
                                    ProjF = ProjF + "\\";
                                }

                                string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;

                                Append_ML_object_data();
                                dataGridView_sheet_index.DataSource = _SGEN_mainform.dt_sheet_index;
                                dataGridView_sheet_index.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                dataGridView_sheet_index.DefaultCellStyle.ForeColor = Color.White;
                                dataGridView_sheet_index.EnableHeadersVisualStyles = false;


                                label_not_saved.Visible = true;
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            set_enable_true();

            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void Populate_data_table_matchline_file_names()
        {
            if (_SGEN_mainform.dt_sheet_index != null)
            {
                if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    string No_start = _SGEN_mainform.tpage_settings.get_start_number_from_text_box();
                    string Preffix = _SGEN_mainform.tpage_settings.get_prefix_name_from_text_box();
                    string Suffix = _SGEN_mainform.tpage_settings.get_suffix_name_from_text_box();

                    int Increment = 1;
                    if (Functions.IsNumeric(_SGEN_mainform.tpage_settings.get_increment_from_text_box()) == true)
                    {
                        Increment = Convert.ToInt32(_SGEN_mainform.tpage_settings.get_increment_from_text_box());
                    }


                    if (Functions.IsNumeric(No_start) == true)
                    {
                        int nr_start = Convert.ToInt32(No_start);
                        int old_nr = nr_start;

                        for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            string new_nr = old_nr.ToString();
                            if (i > 0) new_nr = (old_nr + Increment).ToString();
                            int len_no_start = No_start.Length;
                            int Len_new = new_nr.Length;
                            if (len_no_start > Len_new)
                            {
                                for (int j = Len_new; j < len_no_start; ++j)
                                {
                                    new_nr = "0" + new_nr;
                                }
                            }
                            string File_name = Preffix + new_nr + Suffix;
                            _SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] = File_name;
                            old_nr = Convert.ToInt32(new_nr);
                        }
                    }
                }
            }

        }

        private Polyline create_rectangle_VP(double scale1, Point3d Point1, string anchor, int cid)
        {



            Polyline poly1 = new Autodesk.AutoCAD.DatabaseServices.Polyline();

            if (anchor == "TL")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
            }

            if (anchor == "TR")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X, Point1.Y - _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
            }

            if (anchor == "BR")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width / scale1, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
            }

            if (anchor == "BL")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X, Point1.Y), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y + _SGEN_mainform.Vw_height / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width / scale1, Point1.Y), 0, 0, 0);

            }

            if (anchor == "CEN")
            {
                poly1.AddVertexAt(0, new Point2d(Point1.X - _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y - _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(1, new Point2d(Point1.X - _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y + _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(2, new Point2d(Point1.X + _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y + _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);
                poly1.AddVertexAt(3, new Point2d(Point1.X + _SGEN_mainform.Vw_width * 0.5 / scale1, Point1.Y - _SGEN_mainform.Vw_height * 0.5 / scale1), 0, 0, 0);

            }

            poly1.Closed = true;
            poly1.ColorIndex = cid;
            poly1.Elevation = 0;

            return poly1;
        }


        private void Create_ML_object_data()
        {

            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForWrite) as BlockTable;

                        using (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord)
                        {
                            if (BTrecord != null)
                            {
                                List<string> List1 = new List<string>();
                                List<string> List2 = new List<string>();
                                List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();

                                List1.Add("MMID");
                                List2.Add("ObjectID of the rectangle");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("DrawingNum");
                                List2.Add("Alignment_number");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("BeginSta");
                                List2.Add("Matchline start");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("EndSta");
                                List2.Add("Matchline end");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Center_X");
                                List2.Add("X in modelspace");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Center_Y");
                                List2.Add("Y in modelspace");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Rotation");
                                List2.Add("E-W viewport line rotation");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Width");
                                List2.Add("Matchline rectangle width");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Height");
                                List2.Add("Matchline rectangle height");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                List1.Add("Type");
                                List2.Add("Type of drawing related to the rectangle");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("Note1");
                                List2.Add("Notes");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("Version");
                                List2.Add("Version number");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("DateMod");
                                List2.Add("DateMod");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                List1.Add("SegmentName");
                                List2.Add("SegmentName");
                                List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Functions.Get_object_data_table(_SGEN_mainform.od_table_sheet_index, "Generated by SGEN", List1, List2, List3);
                            }
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

        private void button1_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForWrite) as BlockTable;



                        List<string> lista_nume = new List<string>();
                        lista_nume.Add("A");
                        lista_nume.Add("B");

                        List<Autodesk.Gis.Map.Constants.DataType> lista_types = new List<Autodesk.Gis.Map.Constants.DataType>();
                        lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                        lista_types.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                        Functions.Get_object_data_table("test", "descr", lista_nume, lista_nume, lista_types);

                        Trans1.Commit();

                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();



        }

        private void Create_ML_object_dataTABLE()
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                    if (Tables1.IsTableDefined(_SGEN_mainform.od_table_sheet_index) == true)
                    {
                        Trans1.Dispose();
                        return;
                    }

                    MapApplication app = HostMapApplicationServices.Application;
                    FieldDefinitions tabDefs = app.ActiveProject.MapUtility.NewODFieldDefinitions();



                    List<string> List1 = new List<string>();
                    List<string> List2 = new List<string>();
                    List<Autodesk.Gis.Map.Constants.DataType> List3 = new List<Autodesk.Gis.Map.Constants.DataType>();


                    List1.Add("MMID");
                    List2.Add("ObjectID of the rectangle");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("DrawingNum");
                    List2.Add("Alignment_number");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("BeginSta");
                    List2.Add("Matchline start");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("EndSta");
                    List2.Add("Matchline end");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Center_X");
                    List2.Add("X in modelspace");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Center_Y");
                    List2.Add("Y in modelspace");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Rotation");
                    List2.Add("E-W viewport line rotation");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Width");
                    List2.Add("Matchline rectangle width");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Height");
                    List2.Add("Matchline rectangle height");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                    List1.Add("Type");
                    List2.Add("Type of drawing related to the rectangle");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("Note1");
                    List2.Add("Notes");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("Version");
                    List2.Add("Version number");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("DateMod");
                    List2.Add("DateMod");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                    List1.Add("SegmentName");
                    List2.Add("SegmentName");
                    List3.Add(Autodesk.Gis.Map.Constants.DataType.Character);


                    for (int i = List1.Count - 1; i >= 0; --i)
                    {
                        FieldDefinition def1 = FieldDefinition.Create(List1[i], List2[i], List3[i]);
                        tabDefs.AddColumn(def1, 0);
                    }

                    Tables1.Add(_SGEN_mainform.od_table_sheet_index, tabDefs, "SGEN", true);

                    Trans1.Commit();
                }

            }






        }

        private void button_copy_sheet_index_Click(object sender, EventArgs e)
        {
            if (_SGEN_mainform.no_of_segments > 1)
            {
                if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                {
                    if (_SGEN_mainform.tpage_settings.check_combobox_segment_is_first_one() == true)
                    {

                        if (_SGEN_mainform.dt_sheet_index != null && _SGEN_mainform.dt_sheet_index.Rows.Count > 1)
                        {
                            if (_SGEN_mainform.dt_segments != null && _SGEN_mainform.dt_segments.Rows.Count > 1)
                            {
                                for (int i = 1; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                                {
                                    string ProjF = _SGEN_mainform.project_main_folder;
                                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                    {
                                        ProjF = ProjF + "\\";
                                    }


                                    if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value && _SGEN_mainform.dt_segments.Rows[i]["Start numbering"] != DBNull.Value)
                                    {
                                        string segment1 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);
                                        if (segment1 != "")
                                        {
                                            ProjF = ProjF + segment1 + "\\";
                                        }

                                        string fisier_si = ProjF + _SGEN_mainform.sheet_index_excel_name;
                                        Functions.create_backup(fisier_si);

                                        System.Data.DataTable dt1 = _SGEN_mainform.dt_sheet_index.Copy();
                                        string No_start = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Start numbering"]);
                                        string Preffix = "";
                                        string Suffix = "";
                                        int Increment = 1;

                                        if (_SGEN_mainform.dt_segments.Rows[i]["Prefix File Name"] != DBNull.Value)
                                        {
                                            Preffix = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Prefix File Name"]);
                                        }

                                        if (_SGEN_mainform.dt_segments.Rows[i]["Suffix File Name"] != DBNull.Value)
                                        {
                                            Suffix = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Suffix File Name"]);
                                        }

                                        if (_SGEN_mainform.dt_segments.Rows[i]["Increment"] != DBNull.Value)
                                        {
                                            string val = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Increment"]);
                                            if (Functions.IsNumeric(val) == true)
                                            {
                                                Increment = Convert.ToInt32(val);
                                            }
                                        }

                                        Populate_data_table_matchline_file_names_from_segment_1(dt1, No_start, Preffix, Suffix, Increment);

                                        Populate_sheet_index_file_from_segment_1(fisier_si, dt1);
                                    }

                                }
                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show("please switch to the first segment!", "SGEN", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }


                }
            }
            else
            {
                MessageBox.Show("NO SEGMENTS!", "SGEN", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        public void Populate_sheet_index_file_from_segment_1(string File1, System.Data.DataTable dt1)
        {
            try
            {

                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                if (System.IO.File.Exists(File1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }

                else
                {
                    Workbook1 = Excel1.Workbooks.Open(File1);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    string segment1 = "";

                    Functions.Transfer_to_worksheet_Data_table(W1, dt1, _SGEN_mainform.Start_row_Sheet_index, "General");
                    Create_header_sheet_index_file(W1, _SGEN_mainform.tpage_settings.Get_client_name(), _SGEN_mainform.tpage_settings.Get_project_name(), segment1);

                    W1.Range["A:A"].ColumnWidth = 15;
                    W1.Range["B:C"].ColumnWidth = 20;
                    W1.Range["D:H"].ColumnWidth = 2;
                    W1.Range["I:M"].ColumnWidth = 15;
                    W1.Range["N:Q"].ColumnWidth = 2;
                    W1.Name = "SI_" + System.DateTime.Now.Year.ToString() + "_" + System.DateTime.Now.Month.ToString() + "_" + System.DateTime.Now.Day.ToString();
                    if (System.IO.File.Exists(File1) == false)
                    {
                        Workbook1.SaveAs(File1);
                    }
                    else
                    {
                        Workbook1.Save();
                    }
                    Workbook1.Close();
                    if (Excel1.Workbooks.Count == 0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }
                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void Populate_data_table_matchline_file_names_from_segment_1(System.Data.DataTable dt1, string No_start, string Preffix, string Suffix, int Increment)
        {
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    if (Functions.IsNumeric(No_start) == true)
                    {
                        int nr_start = Convert.ToInt32(No_start);
                        int old_nr = nr_start;

                        for (int i = 0; i < dt1.Rows.Count; ++i)
                        {
                            string new_nr = old_nr.ToString();
                            if (i > 0) new_nr = (old_nr + Increment).ToString();
                            int len_no_start = No_start.Length;
                            int Len_new = new_nr.Length;
                            if (len_no_start > Len_new)
                            {
                                for (int j = Len_new; j < len_no_start; ++j)
                                {
                                    new_nr = "0" + new_nr;
                                }
                            }
                            string File_name = Preffix + new_nr + Suffix;
                            dt1.Rows[i][_SGEN_mainform.Col_dwg_name] = File_name;
                            old_nr = Convert.ToInt32(new_nr);
                        }
                    }
                }
            }

        }

        public void make_labels_visible()
        {

        }

        private void button_load_od_field_to_combobox_Click(object sender, EventArgs e)
        {
            if (checkBox_pick_name_from_OD.Checked==false)
            {
                comboBox_od_field.Items.Clear();
                return;
            }
            set_enable_false();
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        comboBox_od_field.Items.Clear();

                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                        Nume_tables = Tables1.GetTableNames();
                        comboBox_od_field.Items.Clear();
                        for (int i = 0; i < Nume_tables.Count; ++i)
                        {
                            string Tabla1 = Nume_tables[i];

                            Functions.add_OD_fieds_to_combobox(Tabla1, comboBox_od_field);
                        }
                        this.Refresh();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }

        private void checkBox_plat_mode_CheckedChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox check1 = sender as System.Windows.Forms.CheckBox;
            if (check1.Checked == true)
            {
                panel_dan.Visible = true;
                //_SGEN_mainform.tpage_scales.Show();
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is scales_form)
                    {
                        Forma1.Focus();
                        Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                        forma3 = Forma1 as scales_form;
                        return;
                    }
                }
                try
                {
                    scales_form forma2 = new scales_form();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                    forma3 = forma2;
                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }


            }
            else
            {
                panel_dan.Visible = false;
                //_SGEN_mainform.tpage_scales.Hide();
            }
        }
    }
}
