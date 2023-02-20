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
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class AGEN_TBLK_Attributes : Form
    {

        _SGEN_mainform Ag = null;

        System.Data.DataTable Display_dataTable = null;
        System.Data.DataTable dt_atr = null;

        List<string> drawing_list = null;

        private ContextMenuStrip ContextMenuStrip_open_alignment;

        public string Excel_tblk_atr = "";
        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_compare);
            lista_butoane.Add(button_dwg_to_excel);
            lista_butoane.Add(button_excel_to_dwg);
            lista_butoane.Add(button_load_block_attributes_to_excel);
            lista_butoane.Add(button_new_excel_file);
            lista_butoane.Add(button_open_excel_tblk_attrib);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_select_excel_file);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_compare);
            lista_butoane.Add(button_dwg_to_excel);
            lista_butoane.Add(button_excel_to_dwg);
            lista_butoane.Add(button_load_block_attributes_to_excel);
            lista_butoane.Add(button_new_excel_file);
            lista_butoane.Add(button_open_excel_tblk_attrib);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_select_excel_file);



            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public AGEN_TBLK_Attributes()
        {
            InitializeComponent();
            label_block_attributes.Text = "";
            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Open selected drawing" };
            toolStripMenuItem1.Click += open_DWG_Click;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Add drawings" };
            toolStripMenuItem2.Click += button_select_drawings_Click;


            var toolStripMenuItem3 = new ToolStripMenuItem { Text = "Remove drawing" };
            toolStripMenuItem3.Click += remove_selected_dwg_Click;

            var toolStripMenuItem4 = new ToolStripMenuItem { Text = "Clear drawing list" };
            toolStripMenuItem4.Click += remove_all_dwg_Click;

            ContextMenuStrip_open_alignment = new ContextMenuStrip();
            ContextMenuStrip_open_alignment.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2, toolStripMenuItem3, toolStripMenuItem4 });

        }



        private void remove_selected_dwg_Click(object sender, EventArgs e)
        {
            if (dataGridView_drawings.RowCount > 0)
            {
                int Index1 = dataGridView_drawings.CurrentCell.RowIndex;
                if (Index1 == -1)
                {
                    return;
                }

                string val1 = Convert.ToString(dataGridView_drawings.Rows[Index1].Cells[0].Value);
                dataGridView_drawings.Rows.RemoveAt(Index1);
                if (drawing_list != null)
                {
                    if (drawing_list.Contains(val1) == true)
                    {
                        drawing_list.Remove(val1);
                    }
                }
            }
        }

        private void remove_all_dwg_Click(object sender, EventArgs e)
        {
            Display_dataTable = null;

            dataGridView_drawings.DataSource = "";
            drawing_list = new List<string>();
        }

        private void open_DWG_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_drawings.RowCount > 0)
                {

                    int Index1 = dataGridView_drawings.CurrentCell.RowIndex;
                    if (Display_dataTable != null)
                    {
                        if (Display_dataTable.Rows.Count - 1 >= Index1)
                        {
                            string fisier_generat = Display_dataTable.Rows[Index1][0].ToString();

                            string Output_folder = _SGEN_mainform.tpage_settings.get_output_folder_from_text_box();
                            if (Output_folder.Length == 0) Output_folder = "nodata";

                            if (Output_folder.Substring(Output_folder.Length - 1, 1) != "\\")
                            {
                                Output_folder = Output_folder + "\\";
                            }

                            string path1 = Output_folder + fisier_generat + ".dwg";
                            string path2 = fisier_generat;
                            string path0 = "";

                            if (System.IO.File.Exists(path1) == true)
                            {
                                path0 = path1;
                            }

                            if (System.IO.File.Exists(path2) == true)
                            {
                                path0 = path2;
                            }

                            if (System.IO.File.Exists(path0) == true)
                            {

                                bool is_opened = false;
                                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                                foreach (Document Doc in DocumentManager1)
                                {
                                    if (Doc.Name == path0)
                                    {
                                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Doc;
                                        is_opened = true;

                                    }

                                }

                                if (is_opened == false)
                                {
                                    DocumentCollectionExtension.Open(DocumentManager1, path0, false);
                                }

                            }
                            else
                            {
                                MessageBox.Show("file not found");
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

        private void dataGridView_drawings_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_drawings.CurrentCell = dataGridView_drawings.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_open_alignment.Show(Cursor.Position);
                ContextMenuStrip_open_alignment.Visible = true;
            }
            else
            {
                ContextMenuStrip_open_alignment.Visible = false;
            }
        }


        private void dataGridView_drawings_Click(object sender, EventArgs e)
        {
            Type t = e.GetType();
            if (t.Equals(typeof(MouseEventArgs)))
            {
                MouseEventArgs mouse = (MouseEventArgs)e;
                if (mouse.Button == MouseButtons.Right)
                {
                    ContextMenuStrip_open_alignment.Show(Cursor.Position);
                    ContextMenuStrip_open_alignment.Visible = true;
                }
            }
            else
            {
                ContextMenuStrip_open_alignment.Visible = false;
            }
        }
        private void button_dwg_to_excel_Click(object sender, EventArgs e)
        {


            if (MessageBox.Show("This will overwrite tblk attribute excel\r\nDo you want to continue?", "Agen", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;




            try
            {
                set_enable_false();
                _SGEN_mainform Ag = this.MdiParent as _SGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                        {
                            string ProjF = _SGEN_mainform.project_main_folder;
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _SGEN_mainform.block_attributes_excel_name;

                            if (System.IO.File.Exists(fisier_attributes) == true)

                            {
                                file_de_procesat = fisier_attributes;

                            }
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(Excel_tblk_atr) == true)
                        {
                            file_de_procesat = Excel_tblk_atr;
                        }
                    }



                    if (System.IO.File.Exists(file_de_procesat) == true)
                    {

                        Functions.create_backup(file_de_procesat);
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
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(file_de_procesat);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                        try
                        {

                            System.Data.DataTable dt1 = creaza_data_table_attributes(W1);
                            if (dt1 != null)
                            {
                                string segment1 = "";
                                Create_header_block_attributes_file(W1, _SGEN_mainform.tpage_settings.Get_client_name(), "", segment1, dt1.Columns.Count);
                                System.Data.DataTable dt_header = creaza_data_table_excel_header_values(dt1);
                                Transfera_data_to_excel_fara_header(W1, dt_header, _SGEN_mainform.Start_row_block_attributes);
                                Transfer_to_w1_Data_table_values(W1, dt1, _SGEN_mainform.Start_row_block_attributes + 2);
                                Workbook1.Save();
                            }

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
                    else
                    {
                        MessageBox.Show("no tblk_attributes_found", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }


            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            set_enable_true();

        }


        private void Transfer_to_w1_Data_table_values(Microsoft.Office.Interop.Excel.Worksheet W1, System.Data.DataTable Data_table, int Start_row)
        {
            if (Data_table != null)
            {
                if (Data_table.Rows.Count > 0)
                {
                    int NrR = Data_table.Rows.Count;
                    int NrC = Data_table.Columns.Count;

                    Object[,] values = new object[NrR, NrC];
                    for (int i = 0; i < NrR; ++i)
                    {
                        for (int j = 0; j < NrC; ++j)
                        {
                            if (Data_table.Rows[i][j] != DBNull.Value)
                            {
                                string val = Data_table.Rows[i][j].ToString();
                                values[i, j] = val;
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A" + Start_row.ToString() + ":" + Functions.get_excel_column_letter(NrC) + (NrR + Start_row - 1).ToString()];
                    range1.Cells.NumberFormat = "@";
                    range1.Value2 = values;


                    Functions.Color_border_range_inside(range1, 0);




                }
            }
        }

        private System.Data.DataTable creaza_data_table_excel_header_values(System.Data.DataTable dt_val)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < dt_val.Columns.Count; ++i)
            {
                dt1.Columns.Add("col" + i.ToString(), typeof(string));
            }

            dt1.Rows.Add();
            dt1.Rows[0][0] = "***";
            dt1.Rows[0][1] = "Name of Block:";
            dt1.Rows.Add();
            dt1.Rows[1][0] = "Dwg name";
            dt1.Rows[1][1] = "Layout name";

            for (int j = 2; j < dt_val.Columns.Count; ++j)
            {

                string colname = dt_val.Columns[j].ColumnName;
                char split1 = Convert.ToChar("|");

                string[] bl_atr = colname.Split(split1);
                dt1.Rows[0][j] = bl_atr[0];
                dt1.Rows[1][j] = bl_atr[1];

            }


            return dt1;
        }

        private void Transfera_data_to_excel_fara_header(Microsoft.Office.Interop.Excel.Worksheet W1, System.Data.DataTable Data_table, int Start_row)
        {
            if (Data_table != null)
            {
                if (Data_table.Rows.Count > 0)
                {
                    int NrR = Data_table.Rows.Count;
                    int NrC = Data_table.Columns.Count;

                    Object[,] values = new object[NrR, NrC];
                    for (int i = 0; i < NrR; ++i)
                    {
                        for (int j = 0; j < NrC; ++j)
                        {
                            if (Data_table.Rows[i][j] != DBNull.Value)
                            {
                                string val = Data_table.Rows[i][j].ToString();
                                values[i, j] = val;
                            }
                        }
                    }

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A" + Start_row.ToString() + ":" + Functions.get_excel_column_letter(NrC) + (NrR + Start_row - 1).ToString()];
                    range1.Cells.NumberFormat = "@";
                    range1.Value2 = values;

                    if (NrR == 1)
                    {
                        Functions.Color_border_range_inside(range1, 0);
                    }



                }
            }
        }

        private System.Data.DataTable creaza_data_table_attributes(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();



            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A7:XX8"];

            object[,] matrix1 = new object[2, 648];
            matrix1 = range1.Value2;

            dt1.Columns.Add("Drawing", typeof(string));
            dt1.Columns.Add("Layout", typeof(string));

            Display_dataTable = new System.Data.DataTable();
            Display_dataTable.Columns.Add("Drawing", typeof(string));
            Display_dataTable.Columns.Add("Layout", typeof(string));
            Display_dataTable.Columns.Add("Blocks Found", typeof(string));

            int end1 = 0;
            for (int i = 1; i <= 648; ++i)
            {
                object val1 = matrix1[1, i];
                if (val1 == null)
                {
                    end1 = i - 1;
                    i = 649;
                }
            }

            if (end1 < 3)
            {
                MessageBox.Show("no header found for the block attributes file\r\nOperation aborted");
                return null;
            }

            for (int i = 3; i <= end1; ++i)
            {
                dt1.Columns.Add(matrix1[1, i].ToString() + "|" + matrix1[2, i].ToString(), typeof(string));
            }

            string Output_folder = _SGEN_mainform.tpage_settings.get_output_folder_from_text_box();

            if (Output_folder.Length == 0) Output_folder = "nodata";

            if (Output_folder.Substring(Output_folder.Length - 1, 1) != "\\")
            {
                Output_folder = Output_folder + "\\";
            }

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (ThisDrawing == null)
            {
                MessageBox.Show("you are trying to run outside of a drawing\r\nopen or create a drawing\r\noperation aborted");
                set_enable_true();
                return null;
            }

            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                if (drawing_list == null)
                {
                    if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        drawing_list = new List<string>();
                        for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name].ToString();
                                drawing_list.Add(file1);
                            }
                        }
                    }
                }

                if (drawing_list.Count == 0)
                {
                    if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name].ToString();
                                drawing_list.Add(file1);
                            }
                        }
                    }
                }

                if (drawing_list != null)
                {
                    if (drawing_list.Count > 0)
                    {
                        for (int i = 0; i < drawing_list.Count; ++i)
                        {
                            string file1 = drawing_list[i];
                            string path0 = "";
                            string path1 = Output_folder + file1 + ".dwg";
                            string path2 = file1;
                            if (System.IO.File.Exists(path1) == true)
                            {
                                path0 = path1;
                            }
                            if (System.IO.File.Exists(path2) == true)
                            {
                                path0 = path2;
                            }

                            if (System.IO.File.Exists(path0) == true)
                            {
                                using (Database Database2 = new Database(false, true))
                                {
                                    Database2.ReadDwgFile(path0, FileOpenMode.OpenForReadAndAllShare, true, "");
                                    //System.IO.FileShare.ReadWrite, false, null);
                                    Database2.CloseInput(true);
                                    HostApplicationServices.WorkingDatabase = Database2;
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                    {
                                        DBDictionary Layoutdict = (DBDictionary)Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead);

                                        LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

                                        Layout Layout1 = null;
                                        foreach (DBDictionaryEntry entry in Layoutdict)
                                        {
                                            Layout1 = (Layout)Trans2.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                                            if (Layout1.TabOrder > 0)
                                            {
                                                //LayoutManager1.CurrentLayout = Layout1.LayoutName;

                                                dt1.Rows.Add();
                                                Display_dataTable.Rows.Add();
                                                string found1 = "NO";

                                                BlockTableRecord BtrecordPS = (BlockTableRecord)Trans2.GetObject(Layout1.BlockTableRecordId, OpenMode.ForRead);

                                                foreach (ObjectId id1 in BtrecordPS)
                                                {
                                                    BlockReference block1 = Trans2.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                                    if (block1 != null)
                                                    {
                                                        if (block1.AttributeCollection.Count > 0)
                                                        {
                                                            string block_name = Functions.get_block_name_another_database(block1, Database2);


                                                            for (int j = 2; j < dt1.Columns.Count; ++j)
                                                            {

                                                                string colname = dt1.Columns[j].ColumnName;
                                                                char split1 = Convert.ToChar("|");

                                                                string[] bl_atr = colname.Split(split1);
                                                                if (block_name.ToLower() == bl_atr[0].ToLower())
                                                                {
                                                                    found1 = "YES";

                                                                    foreach (ObjectId atid in block1.AttributeCollection)
                                                                    {
                                                                        AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForRead);

                                                                        if (Atr1 != null)
                                                                        {
                                                                            if (Atr1.Tag.ToLower() == bl_atr[1].ToLower() && Atr1.TextString != "")
                                                                            {
                                                                                if (Atr1.IsMTextAttribute == false)
                                                                                {
                                                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.TextString;
                                                                                }
                                                                                else
                                                                                {
                                                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.MTextAttribute.Contents;
                                                                                }

                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Drawing"] = file1;
                                                dt1.Rows[dt1.Rows.Count - 1]["Layout"] = Layout1.LayoutName;

                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Drawing"] = file1;
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Layout"] = Layout1.LayoutName;
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Blocks Found"] = found1;
                                            }
                                        }
                                    }

                                    HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                }
                            }

                        }
                    }
                }
            }


            //   Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);

            dataGridView_drawings.DataSource = Display_dataTable;
            dataGridView_drawings.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_drawings.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_drawings.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_drawings.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_drawings.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_drawings.EnableHeadersVisualStyles = false;

            return dt1;
        }





        private void button_excel_to_dwg_Click(object sender, EventArgs e)
        {


            Display_dataTable = new System.Data.DataTable();
            Display_dataTable.Columns.Add("Drawing", typeof(string));
            Display_dataTable.Columns.Add("Layout", typeof(string));
            Display_dataTable.Columns.Add("Blocks Populated", typeof(string));





            try
            {
                set_enable_false();
                _SGEN_mainform Ag = this.MdiParent as _SGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                        {
                            string ProjF = _SGEN_mainform.project_main_folder;
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _SGEN_mainform.block_attributes_excel_name;

                            if (System.IO.File.Exists(fisier_attributes) == true)

                            {
                                file_de_procesat = fisier_attributes;

                            }
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(Excel_tblk_atr) == true)
                        {
                            file_de_procesat = Excel_tblk_atr;
                        }
                    }

                    file_de_procesat = file_de_procesat.Replace("\\\\mottmac.group.int\\Project\\MMNA\\Talon\\Pipeline\\", "G:\\");


                    bool excel_is_opened = false;

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            string file_opened = Workbook2.FullName;
                            file_opened = file_opened.Replace("\\\\mottmac.group.int\\Project\\MMNA\\Talon\\Pipeline\\", "G:\\");
                            if (file_opened == file_de_procesat)
                            {
                                Workbook1 = Workbook2;
                                W1 = Workbook1.Worksheets[1];
                                excel_is_opened = true;
                            }

                        }


                    }
                    catch (System.Exception)
                    {

                    }

                    if (System.IO.File.Exists(file_de_procesat) == true)
                    {

                        if (W1 == null)
                        {
                            try
                            {
                                Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                            }
                            catch (System.Exception ex)
                            {
                                Excel1 = new Microsoft.Office.Interop.Excel.Application();

                            }

                            if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                            Workbook1 = Excel1.Workbooks.Open(file_de_procesat);
                            W1 = Workbook1.Worksheets[1];
                        }



                        try
                        {

                            System.Data.DataTable dt1 = Load_attributes_from_excel(W1);

                            if (excel_is_opened == false)
                            {
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




                            if (dt1 != null)
                            {
                                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                                if (ThisDrawing == null)
                                {
                                    MessageBox.Show("you are trying to run outside of a drawing\r\nopen or create a drawing\r\noperation aborted");
                                    set_enable_true();
                                    return;
                                }
                                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                                {
                                    if (dt1.Rows.Count > 2)
                                    {
                                        if (drawing_list != null)
                                        {
                                            for (int i = dt1.Rows.Count - 1; i >= 2; --i)
                                            {
                                                if (dt1.Rows[i]["Drawing"] != DBNull.Value)
                                                {
                                                    string file1 = Convert.ToString(dt1.Rows[i]["Drawing"]);

                                                    if (file1.Length > 4 && file1.Substring(file1.Length - 4, 4).ToLower() != ".dwg")
                                                    {
                                                        file1 = file1 + ".dwg";
                                                    }

                                                    bool found = false;
                                                    for (int j = 0; j < drawing_list.Count; ++j)
                                                    {
                                                        if (drawing_list[j].Contains(file1) == true)
                                                        {
                                                            found = true;
                                                            j = drawing_list.Count;
                                                        }
                                                    }

                                                    if (found == false)
                                                    {
                                                        dt1.Rows.RemoveAt(i);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (dt1.Rows.Count > 2)
                                    {
                                        for (int i = 2; i < dt1.Rows.Count; ++i)
                                        {
                                            if (dt1.Rows[i]["Drawing"] != DBNull.Value && dt1.Rows[i]["Layout"] != DBNull.Value)
                                            {
                                                string Layout_Excel = dt1.Rows[i]["Layout"].ToString();
                                                string file1 = dt1.Rows[i]["Drawing"].ToString();
                                                if (file1.Length > 4 && file1.Substring(file1.Length - 4, 4).ToLower() != ".dwg")
                                                {
                                                    file1 = file1 + ".dwg";
                                                }
                                                for (int j = 0; j < drawing_list.Count; ++j)
                                                {
                                                    if (drawing_list[j].Contains(file1) == true)
                                                    {
                                                        file1 = drawing_list[j];
                                                        j = drawing_list.Count;
                                                    }
                                                }




                                                if (System.IO.File.Exists(file1) == true)
                                                {

                                                    bool is_opened = false;
                                                    DocumentCollection document_collection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                                                    foreach (Document opened_dwg in document_collection)
                                                    {
                                                        if (opened_dwg.Database.Filename == file1)
                                                        {
                                                            HostApplicationServices.WorkingDatabase = opened_dwg.Database;
                                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans3 = opened_dwg.TransactionManager.StartTransaction())
                                                            {
                                                                populate_attributes_from_dt1(ThisDrawing, i, Layout_Excel, Trans3, opened_dwg.Database, dt1, file1, null, null, null);
                                                                is_opened = true;
                                                            }
                                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                        }
                                                    }

                                                    if (is_opened == false)
                                                    {
                                                        using (Database Database2 = new Database(false, true))
                                                        {
                                                            Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndAllShare, true, "");
                                                            //System.IO.FileShare.ReadWrite, false, null);
                                                            Database2.CloseInput(true);
                                                            HostApplicationServices.WorkingDatabase = Database2;
                                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                                            {
                                                                populate_attributes_from_dt1(ThisDrawing, i, Layout_Excel, Trans2, Database2, dt1, file1, Excel1, Workbook1, W1);
                                                            }

                                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                            Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                        }
                                                    }


                                                }

                                                else
                                                {
                                                    MessageBox.Show("the file " + file1 + " was not found\r\noperation aborted");
                                                    set_enable_true();
                                                    return;
                                                }
                                            }
                                        }

                                        MessageBox.Show("done");
                                        dataGridView_drawings.DataSource = Display_dataTable;
                                        dataGridView_drawings.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                        dataGridView_drawings.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                        dataGridView_drawings.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                        dataGridView_drawings.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                        dataGridView_drawings.DefaultCellStyle.ForeColor = Color.White;
                                        dataGridView_drawings.EnableHeadersVisualStyles = false;
                                    }
                                    else
                                    {
                                        MessageBox.Show("nothing updated");
                                    }
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
                    }
                    else
                    {
                        dataGridView_drawings.DataSource = null;
                        Display_dataTable = new System.Data.DataTable();
                        drawing_list = new List<string>();
                        MessageBox.Show("you do not have the TitleBlock attributes file", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                dataGridView_drawings.DataSource = null;
                drawing_list = new List<string>();
                Display_dataTable = new System.Data.DataTable();
            }
            set_enable_true();

        }

        private void populate_attributes_from_dt1(Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing,
                                                int i, string Layout_Excel, Transaction Trans2, Database Database2, System.Data.DataTable dt1, string file1,
                                                    Microsoft.Office.Interop.Excel.Application Excel1, Microsoft.Office.Interop.Excel.Workbook Workbook1, Microsoft.Office.Interop.Excel.Worksheet W1)
        {

            if (Layout_Excel.ToUpper() != "MODEL")
            {
                #region Paper space
                DBDictionary Layoutdict = (DBDictionary)Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead);

                Layout Layout1 = null;
                LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

                foreach (DBDictionaryEntry entry in Layoutdict)
                {
                    Layout Layout0 = (Layout)Trans2.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                    if (Layout0.LayoutName.ToLower() == Layout_Excel.ToLower())
                    {
                        Layout1 = Layout0;
                    }
                }

                string found1 = "NO";
                Display_dataTable.Rows.Add();
                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Drawing"] = file1;

                if (Layout1 != null)
                {
                    //LayoutManager1.CurrentLayout = Layout1.LayoutName;

                    Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Layout"] = Layout1.LayoutName;

                    BlockTableRecord BtrecordPS = (BlockTableRecord)Trans2.GetObject(Layout1.BlockTableRecordId, OpenMode.ForWrite);

                    foreach (ObjectId id1 in BtrecordPS)
                    {
                        BlockReference block1 = Trans2.GetObject(id1, OpenMode.ForRead) as BlockReference;
                        if (block1 != null)
                        {
                            string block_name = Functions.get_block_name_another_database(block1, Database2);

                            for (int j = 2; j < dt1.Columns.Count; ++j)
                            {

                                string block_name_excel = Convert.ToString(dt1.Rows[0][j]);
                                string atributte_name_excel = Convert.ToString(dt1.Rows[1][j]);

                                if (block_name_excel != null && atributte_name_excel != null)
                                {
                                    if (block_name.ToLower() == block_name_excel.ToLower())
                                    {
                                        block1.UpgradeOpen();
                                        if (block1.AttributeCollection.Count > 0)
                                        {
                                            foreach (ObjectId atid in block1.AttributeCollection)
                                            {
                                                AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForWrite);

                                                if (Atr1 != null)
                                                {
                                                    if (Atr1.Tag.ToLower() == atributte_name_excel.ToLower())
                                                    {
                                                        if (dt1.Rows[i][j] != DBNull.Value)
                                                        {
                                                            Atr1.TextString = Convert.ToString(dt1.Rows[i][j]);
                                                        }
                                                        else
                                                        {
                                                            Atr1.TextString = "";
                                                        }

                                                        found1 = "YES";
                                                    }
                                                }
                                            }
                                        }

                                        if (atributte_name_excel.ToLower() == "visibility")
                                        {
                                            if (dt1.Rows[i][j] != DBNull.Value)
                                            {
                                                Functions.set_block_visibility(block1, Convert.ToString(dt1.Rows[i][j]));
                                            }

                                        }

                                    }
                                }
                            }
                        }
                    }
                    Trans2.Commit();
                    Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Blocks Populated"] = found1;
                }
                else
                {
                    MessageBox.Show("the layout " + Layout_Excel + " was not found\r\noperation aborted");
                    set_enable_true();
                    HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    return;
                }
            }

            #endregion
            else
            {
                string found1 = "NO";
                Display_dataTable.Rows.Add();
                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Drawing"] = file1;
                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Layout"] = Layout_Excel;
                BlockTable BlockTable1 = Database2.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                BlockTableRecord BtrecordMS = (BlockTableRecord)Trans2.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                foreach (ObjectId id1 in BtrecordMS)
                {
                    BlockReference block1 = Trans2.GetObject(id1, OpenMode.ForRead) as BlockReference;
                    if (block1 != null)
                    {

                        string block_name = Functions.get_block_name_another_database(block1, Database2);

                        for (int j = 2; j < dt1.Columns.Count; ++j)
                        {

                            string block_name_excel = Convert.ToString(dt1.Rows[0][j]);
                            string atributte_name_excel = Convert.ToString(dt1.Rows[1][j]);

                            if (block_name_excel != null && atributte_name_excel != null)
                            {
                                if (block_name.ToLower() == block_name_excel.ToLower())
                                {
                                    block1.UpgradeOpen();
                                    if (block1.AttributeCollection.Count > 0)
                                    {
                                        foreach (ObjectId atid in block1.AttributeCollection)
                                        {
                                            AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForWrite);

                                            if (Atr1 != null)
                                            {
                                                if (Atr1.Tag.ToLower() == atributte_name_excel.ToLower())
                                                {
                                                    if (dt1.Rows[i][j] != DBNull.Value)
                                                    {
                                                        Atr1.TextString = Convert.ToString(dt1.Rows[i][j]);
                                                    }
                                                    else
                                                    {
                                                        Atr1.TextString = "";
                                                    }

                                                    found1 = "YES";
                                                }
                                            }
                                        }

                                    }



                                    if (atributte_name_excel.ToLower() == "visibility")
                                    {
                                        if (dt1.Rows[i][j] != DBNull.Value)
                                        {
                                            Functions.set_block_visibility(block1, Convert.ToString(dt1.Rows[i][j]));
                                        }

                                    }
                                }
                            }
                        }
                    }
                }

                Trans2.Commit();
                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Blocks Populated"] = found1;
            }
        }

        private System.Data.DataTable Load_attributes_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1)
        {


            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Drawing", typeof(string));
            dt1.Columns.Add("Layout", typeof(string));



            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A7:XX8"];
            object[,] matrix1 = new object[2, 648];
            matrix1 = range1.Value2;



            int col_end1 = 2;
            for (int j = 3; j <= 648; ++j)
            {
                object val1 = matrix1[1, j];
                object val2 = matrix1[2, j];

                if (val1 == null || val2 == null)
                {
                    col_end1 = j - 1;
                    j = 649;
                }
            }

            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A1:A30000"];

            object[,] matrix2 = new object[1, 30000];
            matrix2 = range2.Value2;



            int row_end1 = 8;
            for (int i = 9; i <= 30000; ++i)
            {
                object val1 = matrix2[i, 1];

                if (val1 == null)
                {
                    row_end1 = i - 1;
                    i = 30001;
                }
            }

            if (col_end1 < 3)
            {
                MessageBox.Show("no header found for the block attributes file\r\nOperation aborted");
                return null;
            }

            if (row_end1 < 9)
            {
                MessageBox.Show("no attribute data found\r\nOperation aborted");
                return null;
            }

            for (int j = 3; j <= col_end1; ++j)
            {
                dt1.Columns.Add("col_" + j.ToString(), typeof(string));
            }


            Microsoft.Office.Interop.Excel.Range range3 = W1.Range[W1.Cells[1, 1], W1.Cells[row_end1, col_end1]];

            object[,] matrix3 = new object[row_end1, col_end1];
            matrix3 = range3.Value2;

            for (int i = 7; i <= 8; ++i)
            {
                dt1.Rows.Add();

                dt1.Rows[dt1.Rows.Count - 1]["Drawing"] = "***";
                dt1.Rows[dt1.Rows.Count - 1]["Layout"] = "***";

                for (int j = 3; j <= col_end1; ++j)
                {
                    object val1 = matrix3[i, j];
                    if (val1 != null)
                    {
                        dt1.Rows[dt1.Rows.Count - 1][j - 1] = val1.ToString();
                    }
                }

            }

            for (int i = 9; i <= row_end1; ++i)
            {
                dt1.Rows.Add();

                for (int j = 1; j <= col_end1; ++j)
                {
                    object val1 = matrix3[i, j];
                    if (val1 != null)
                    {
                        dt1.Rows[dt1.Rows.Count - 1][j - 1] = val1.ToString();
                    }
                }

            }


            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);


            return dt1;
        }


        private void button_open_excel_tblk_attributes_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {

                string file_de_procesat = "";

                if (Excel_tblk_atr == "")
                {
                    if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                    {
                        string ProjF = _SGEN_mainform.project_main_folder;
                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }

                        string fisier_attributes = ProjF + _SGEN_mainform.block_attributes_excel_name;

                        if (System.IO.File.Exists(fisier_attributes) == true)

                        {
                            file_de_procesat = fisier_attributes;

                        }
                    }
                }
                else
                {
                    if (System.IO.File.Exists(Excel_tblk_atr) == true)
                    {
                        file_de_procesat = Excel_tblk_atr;
                    }
                }




                if (System.IO.File.Exists(file_de_procesat) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the block attributes data file does not exist");
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
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(file_de_procesat);




            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();


        }

        private void button_select_drawings_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "Autocad files (*.dwg)|*.dwg";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (Display_dataTable == null)
                    {
                        Display_dataTable = new System.Data.DataTable();
                        Display_dataTable.Columns.Add("Drawing", typeof(string));
                    }

                    if (drawing_list == null)
                    {
                        drawing_list = new List<string>();
                    }

                    _SGEN_mainform Ag = this.MdiParent as _SGEN_mainform;


                    foreach (string file1 in fbd.FileNames)
                    {
                        if (drawing_list.Contains(file1) == false)
                        {
                            Display_dataTable.Rows.Add();
                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][0] = file1;
                            drawing_list.Add(file1);
                        }
                    }

                    Display_dataTable = Functions.Sort_data_table(Display_dataTable, "Drawing");

                    dataGridView_drawings.DataSource = Display_dataTable;
                    dataGridView_drawings.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dataGridView_drawings.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_drawings.DefaultCellStyle.ForeColor = Color.White;
                    dataGridView_drawings.EnableHeadersVisualStyles = false;

                    Ag.WindowState = FormWindowState.Normal;
                }
            }
        }

        private void button_load_block_attributes_to_excel_header_Click(object sender, EventArgs e)
        {



            if (Functions.Get_if_workbook_is_open_in_Excel(Excel_tblk_atr) == true)
            {
                MessageBox.Show("Please close the " + Excel_tblk_atr + " file");
                return;
            }

            Ag = this.MdiParent as _SGEN_mainform;

            Ag.WindowState = FormWindowState.Minimized;

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (ThisDrawing == null)
            {
                MessageBox.Show("you are trying to run outside of a drawing\r\nopen or create a drawing\r\noperation aborted");
                set_enable_true();
                return;
            }
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Editor1.SetImpliedSelection(Empty_array);
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect block references having attributes:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            set_enable_true();

                            Set_label_block_attributes_to_red();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            Ag.WindowState = FormWindowState.Normal;

                            return;
                        }

                        List<string> Lista_bl = new List<string>();
                        List<string> Lista_at = new List<string>();
                        Lista_at.Add("Dwg name");
                        Lista_bl.Add("***");
                        Lista_at.Add("Layout name");
                        Lista_bl.Add("Name of Block:");


                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            BlockReference Block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (Block1 != null)
                            {
                                if (Block1.AttributeCollection.Count > 0)
                                {
                                    foreach (ObjectId id in Block1.AttributeCollection)
                                    {
                                        if (id.IsErased == false)
                                        {
                                            AttributeReference attRef = Trans1.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as AttributeReference;
                                            if (attRef != null)
                                            {
                                                string Continut = attRef.TextString;
                                                string Tag = attRef.Tag;
                                                Lista_at.Add(Tag);
                                                Lista_bl.Add(Block1.Name);
                                            }
                                        }
                                    }
                                }

                            }
                        }

                        if (Lista_at.Count > 0)
                        {
                            dt_atr = new System.Data.DataTable();

                            for (int i = 0; i < Lista_at.Count; ++i)
                            {
                                dt_atr.Columns.Add(i.ToString(), typeof(string));
                            }
                            dt_atr.Rows.Add();
                            dt_atr.Rows.Add();
                            for (int i = 0; i < Lista_at.Count; ++i)
                            {
                                dt_atr.Rows[0][i] = Lista_bl[i];
                                dt_atr.Rows[1][i] = Lista_at[i];
                            }

                            if (Excel_tblk_atr == "")
                            {
                                if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                                {
                                    string ProjF = _SGEN_mainform.project_main_folder;
                                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                    {
                                        ProjF = ProjF + "\\";
                                    }
                                    string fisier_atr = ProjF + _SGEN_mainform.block_attributes_excel_name;

                                    Populate_block_attributes_file(fisier_atr);
                                    Set_label_block_attributes_to_green();
                                }
                                else
                                {
                                    Set_label_block_attributes_to_red();
                                }
                            }
                            else
                            {
                                if (System.IO.File.Exists(Excel_tblk_atr) == true)
                                {

                                    Populate_block_attributes_file(Excel_tblk_atr);
                                    Set_label_block_attributes_to_green();
                                }
                                else
                                {
                                    Set_label_block_attributes_to_red();
                                }
                            }
                        }
                        else
                        {
                            Set_label_block_attributes_to_red();
                        }
                        Trans1.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                Set_label_block_attributes_to_red();
            }

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();



            Ag.WindowState = FormWindowState.Normal;

        }

        public void Populate_block_attributes_file(String File1)
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
                    Functions.create_backup(File1);
                    Workbook1 = Excel1.Workbooks.Open(File1);
                }

                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    string segment1 = "";
                    Create_header_block_attributes_file(W1, _SGEN_mainform.tpage_settings.Get_client_name(), _SGEN_mainform.tpage_settings.Get_project_name(), segment1, dt_atr.Columns.Count);
                    Transfera_data_to_excel_fara_header(W1, dt_atr, _SGEN_mainform.Start_row_block_attributes);

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

        public void Set_label_block_attributes_to_red()
        {
            label_block_attributes.Text = "Attributes not loaded";
            label_block_attributes.ForeColor = Color.Red;
        }

        public void Set_label_block_attributes_to_green()
        {
            label_block_attributes.Text = "Attributes loaded";
            label_block_attributes.ForeColor = Color.LimeGreen;
        }

        private void button_compare_Click(object sender, EventArgs e)
        {



            if (Functions.Get_if_workbook_is_open_in_Excel(Excel_tblk_atr) == true)
            {
                MessageBox.Show("Please close the " + Excel_tblk_atr + " file");
                return;
            }

            try
            {
                set_enable_false();
                _SGEN_mainform Ag = this.MdiParent as _SGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                        {
                            string ProjF = _SGEN_mainform.project_main_folder;
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _SGEN_mainform.block_attributes_excel_name;

                            if (System.IO.File.Exists(fisier_attributes) == true)

                            {
                                file_de_procesat = fisier_attributes;

                            }
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(Excel_tblk_atr) == true)
                        {
                            file_de_procesat = Excel_tblk_atr;
                        }
                    }



                    if (System.IO.File.Exists(file_de_procesat) == true)
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
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(file_de_procesat);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                        try
                        {
                            System.Data.DataTable dt1 = creaza_data_table_compare(W1);
                            Workbook1.Close();
                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Quit();
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }

                            if (dt1.Rows.Count == 0)
                            {
                                MessageBox.Show("no discrepancies found");
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
                    else
                    {
                        MessageBox.Show("you do not have the TitleBlock attributes file", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }


                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }

        private System.Data.DataTable creaza_data_table_compare(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Drawing", typeof(string));
            dt1.Columns.Add("Layout", typeof(string));


            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A7:XX8"];
            object[,] matrix1 = new object[2, 648];
            matrix1 = range1.Value2;



            int col_end1 = 2;
            for (int j = 3; j <= 648; ++j)
            {
                object val1 = matrix1[1, j];
                object val2 = matrix1[2, j];

                if (val1 == null || val2 == null)
                {
                    col_end1 = j - 1;
                    j = 649;
                }
            }

            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A1:A30000"];

            object[,] matrix2 = new object[1, 30000];
            matrix2 = range2.Value2;



            int row_end1 = 8;
            for (int i = 9; i <= 30000; ++i)
            {
                object val1 = matrix2[i, 1];

                if (val1 == null)
                {
                    row_end1 = i - 1;
                    i = 30001;
                }
            }

            if (col_end1 < 3)
            {
                MessageBox.Show("no header found for the block attributes file\r\nOperation aborted");
                return null;
            }

            if (row_end1 < 9)
            {
                MessageBox.Show("no attribute data found\r\nOperation aborted");
                return null;
            }


            for (int j = 3; j <= col_end1; ++j)
            {
                dt1.Columns.Add(matrix1[1, j].ToString() + "|" + matrix1[2, j].ToString(), typeof(string));
            }

            string Output_folder = _SGEN_mainform.tpage_settings.get_output_folder_from_text_box();

            if (Output_folder.Length == 0) Output_folder = "nodata";

            if (Output_folder.Substring(Output_folder.Length - 1, 1) != "\\")
            {
                Output_folder = Output_folder + "\\";
            }

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (ThisDrawing == null)
            {
                MessageBox.Show("you are trying to run outside of a drawing\r\nopen or create a drawing\r\noperation aborted");
                set_enable_true();
                return null;
            }

            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                if (drawing_list == null)
                {
                    if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        drawing_list = new List<string>();
                        for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name].ToString();
                                drawing_list.Add(file1);
                            }
                        }
                    }
                }

                if (drawing_list != null)
                {
                    if (drawing_list.Count > 0)
                    {
                        for (int i = 0; i < drawing_list.Count; ++i)
                        {
                            string file1 = drawing_list[i];
                            string path0 = "";
                            string path1 = Output_folder + file1 + ".dwg";
                            string path2 = file1;
                            if (System.IO.File.Exists(path1) == true)
                            {
                                path0 = path1;
                            }
                            if (System.IO.File.Exists(path2) == true)
                            {
                                path0 = path2;
                            }

                            if (System.IO.File.Exists(path0) == true)
                            {
                                using (Database Database2 = new Database(false, true))
                                {
                                    Database2.ReadDwgFile(path0, FileOpenMode.OpenForReadAndAllShare, true, "");
                                    //System.IO.FileShare.ReadWrite, false, null);
                                    Database2.CloseInput(true);
                                    HostApplicationServices.WorkingDatabase = Database2;
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                    {
                                        DBDictionary Layoutdict = (DBDictionary)Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead);

                                        LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;

                                        Layout Layout1 = null;
                                        foreach (DBDictionaryEntry entry in Layoutdict)
                                        {
                                            Layout1 = (Layout)Trans2.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                                            if (Layout1.TabOrder > 0)
                                            {
                                                //LayoutManager1.CurrentLayout = Layout1.LayoutName;

                                                dt1.Rows.Add();



                                                BlockTableRecord BtrecordPS = (BlockTableRecord)Trans2.GetObject(Layout1.BlockTableRecordId, OpenMode.ForRead);

                                                foreach (ObjectId id1 in BtrecordPS)
                                                {
                                                    BlockReference block1 = Trans2.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                                    if (block1 != null)
                                                    {
                                                        if (block1.AttributeCollection.Count > 0)
                                                        {
                                                            string block_name = Functions.get_block_name_another_database(block1, Database2);


                                                            for (int j = 2; j < dt1.Columns.Count; ++j)
                                                            {

                                                                string colname = dt1.Columns[j].ColumnName;
                                                                char split1 = Convert.ToChar("|");

                                                                string[] bl_atr = colname.Split(split1);
                                                                if (block_name.ToLower() == bl_atr[0].ToLower())
                                                                {


                                                                    foreach (ObjectId atid in block1.AttributeCollection)
                                                                    {
                                                                        AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForRead);

                                                                        if (Atr1 != null)
                                                                        {
                                                                            if (Atr1.Tag.ToLower() == bl_atr[1].ToLower() && Atr1.TextString != "")
                                                                            {
                                                                                if (Atr1.IsMTextAttribute == false)
                                                                                {
                                                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.TextString;
                                                                                }
                                                                                else
                                                                                {
                                                                                    dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.MTextAttribute.Contents;
                                                                                }

                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                dt1.Rows[dt1.Rows.Count - 1]["Drawing"] = file1;
                                                dt1.Rows[dt1.Rows.Count - 1]["Layout"] = Layout1.LayoutName;


                                            }
                                        }
                                    }

                                    HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                }
                            }

                        }
                    }
                }
            }

            System.Data.DataTable dt3 = new System.Data.DataTable();
            dt3.Columns.Add("Drawing", typeof(string));
            dt3.Columns.Add("Layout", typeof(string));
            dt3.Columns.Add("Block Name", typeof(string));
            dt3.Columns.Add("Attribute Name", typeof(string));
            dt3.Columns.Add("Excel Value", typeof(string));
            dt3.Columns.Add("DWG Value", typeof(string));

            if (dt1.Rows.Count > 0)
            {
                System.Data.DataTable dt2 = Load_attributes_from_excel(W1);

                if (dt2.Rows.Count > 0)
                {

                    for (int i = 0; i < dt1.Rows.Count; ++i)
                    {
                        string file1 = Convert.ToString(dt1.Rows[i]["Drawing"]);
                        string layout1 = Convert.ToString(dt1.Rows[i]["Layout"]);



                        for (int m = 2; m < dt2.Rows.Count; ++m)
                        {
                            string file2 = Convert.ToString(dt2.Rows[m]["Drawing"]);
                            string layout2 = Convert.ToString(dt2.Rows[m]["Layout"]);

                            if (file1 == file2 && layout1 == layout2)
                            {

                                for (int j = 2; j < dt1.Columns.Count; ++j)
                                {

                                    string colname = dt1.Columns[j].ColumnName;
                                    char split1 = Convert.ToChar("|");

                                    string[] bl_atr = colname.Split(split1);
                                    string block_name_dwg = bl_atr[0].ToLower();
                                    string attribute_name_dwg = bl_atr[1].ToLower();

                                    for (int n = 2; n < dt2.Columns.Count; ++n)
                                    {
                                        string block_name_excel = Convert.ToString(dt2.Rows[0][n]).ToLower();
                                        string attributte_name_excel = Convert.ToString(dt2.Rows[1][n]).ToLower();
                                        if (block_name_dwg == block_name_excel && attribute_name_dwg == attributte_name_excel)
                                        {
                                            string val1 = "";
                                            if (dt1.Rows[i][j] != DBNull.Value)
                                            {
                                                val1 = Convert.ToString(dt1.Rows[i][j]);
                                            }
                                            string val2 = "";
                                            if (dt2.Rows[m][n] != DBNull.Value)
                                            {
                                                val2 = Convert.ToString(dt2.Rows[m][n]);
                                            }
                                            if (val1 != val2)
                                            {
                                                dt3.Rows.Add();

                                                dt3.Rows[dt3.Rows.Count - 1]["Drawing"] = file1;
                                                dt3.Rows[dt3.Rows.Count - 1]["Layout"] = layout1;
                                                dt3.Rows[dt3.Rows.Count - 1]["Block Name"] = block_name_dwg;
                                                dt3.Rows[dt3.Rows.Count - 1]["Attribute Name"] = attribute_name_dwg;
                                                dt3.Rows[dt3.Rows.Count - 1]["Excel Value"] = val2;
                                                dt3.Rows[dt3.Rows.Count - 1]["DWG Value"] = val1;
                                            }


                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            Functions.Transfer_datatable_to_new_excel_spreadsheet(dt3);



            return dt3;
        }

        private void button_select_excel_file_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {

                    string File1 = fbd.FileName;
                    label_excel_to_green(System.IO.Path.GetFileNameWithoutExtension(File1));
                    Excel_tblk_atr = File1;

                }
                else
                {
                    label_excel_to_red();
                }
            }
        }


        public void label_excel_to_red()
        {
            label_excel_file.Text = "Excel file not specified";
            label_excel_file.ForeColor = Color.Red;
            Excel_tblk_atr = "";
        }

        private void label_excel_to_green(string file_without_extension)
        {
            label_excel_file.Text = file_without_extension;
            label_excel_file.ForeColor = Color.LimeGreen;
        }







        private void Button_new_excel_file_Click(object sender, EventArgs e)
        {
            if (Functions.Get_if_workbook_is_open_in_Excel(_SGEN_mainform.block_attributes_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _SGEN_mainform.block_attributes_excel_name + " file");
                return;
            }
            string fisier_block_atr = "";
            try
            {
                string ProjFolder = _SGEN_mainform.project_main_folder;
                if (ProjFolder.Length > 0 && ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    fisier_block_atr = ProjFolder + _SGEN_mainform.block_attributes_excel_name;
                    if (System.IO.File.Exists(fisier_block_atr) == true)
                    {
                        MessageBox.Show("there is already a " + _SGEN_mainform.block_attributes_excel_name + " excel file");
                        return;
                    }
                    creaza_new_excel_file(fisier_block_atr);
                }
                else
                {
                    save_new_excel_file();
                }



            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void creaza_new_excel_file(string fisier1)
        {

            if (System.IO.File.Exists(fisier1) == false)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }
                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;


                try
                {

                    Workbook1 = Excel1.Workbooks.Add();
                    Workbook1.SaveAs(fisier1);
                    Workbook1.Close();


                    if (Excel1.Workbooks.Count == 0)
                    {
                        Excel1.Quit();
                    }
                    else
                    {
                        Excel1.Visible = true;
                    }

                    label_excel_to_green(System.IO.Path.GetFileNameWithoutExtension(fisier1));
                    Excel_tblk_atr = fisier1;

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {

                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }




        }

        private void save_new_excel_file()
        {
            SaveFileDialog Save_dlg = new SaveFileDialog();
            Save_dlg.Filter = "Excel file|*.xlsx";
            if (Save_dlg.ShowDialog() == DialogResult.OK)
            {
                string fisier1 = Save_dlg.FileName;
                if (System.IO.File.Exists(fisier1) == false)
                {
                    Microsoft.Office.Interop.Excel.Application Excel1 = null;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();

                    }
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                    try
                    {
                        Workbook1 = Excel1.Workbooks.Add();
                        Workbook1.SaveAs(fisier1);
                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }
                        label_excel_to_green(System.IO.Path.GetFileNameWithoutExtension(fisier1));
                        Excel_tblk_atr = fisier1;
                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }
                }
                else
                {
                    MessageBox.Show("there is already a " + _SGEN_mainform.block_attributes_excel_name + " excel file");
                    return;
                }
            }
        }

        public void Create_header_block_attributes_file(Microsoft.Office.Interop.Excel.Worksheet W1, string Client, string Project, string Segment, int nr_coloane)
        {
            string Last_coloana = Functions.get_excel_column_letter(nr_coloane);

            W1.Columns["A:XX"].Delete();

            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B6"];

            Object[,] valuesH = new object[6, 2];

            valuesH[0, 0] = "CLIENT";
            valuesH[0, 1] = Client;
            valuesH[1, 0] = "PROJECT";
            valuesH[1, 1] = Project;
            valuesH[2, 0] = "SEGMENT";
            valuesH[2, 1] = Segment;
            valuesH[3, 0] = "VERSION";
            valuesH[4, 0] = "DATE CREATED";
            valuesH[4, 1] = DateTime.Now.Year + "-" + DateTime.Now.Month + "-" + DateTime.Now.Day + " at: " + DateTime.Now.TimeOfDay;
            valuesH[5, 0] = "USER ID";
            valuesH[5, 1] = Environment.UserName;



            range1.Value2 = valuesH;

            Functions.Color_border_range_inside(range1, 46);




            Microsoft.Office.Interop.Excel.Range range3 = W1.Range["A7:" + Last_coloana + "7"];
            Functions.Color_border_range_inside(range3, 43);

            Microsoft.Office.Interop.Excel.Range range4 = W1.Range["A8:" + Last_coloana + "8"];
            range4.Font.Color = 16777215;
            range4.Font.Bold = true;
            Functions.Color_border_range_inside(range4, 41);


            Microsoft.Office.Interop.Excel.Range range5 = W1.Range["C1:" + Last_coloana + "6"];
            range5.Merge();
            range5.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            range5.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            range5.Value2 = "TBLK Attributes Table";
            range5.Font.Name = "Arial Black";
            range5.Font.Size = 20;
            range5.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;


            Functions.Color_border_range_outside(range5, 0);


        }
    }
}
