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
using Microsoft.Office.Interop.Excel;

namespace Alignment_mdi
{
    public partial class AGEN_TBLK_Attributes : Form
    {

        _AGEN_mainform Ag = null;

        System.Data.DataTable Display_dataTable = null;
        System.Data.DataTable dt_atr = null;

        List<string> drawing_list = null;

        private ContextMenuStrip ContextMenuStrip_open_alignment;
        private ContextMenuStrip ContextMenuStrip_references;

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
            lista_butoane.Add(button_replace_block);
            lista_butoane.Add(button_load_reference_library);
            lista_butoane.Add(button_select_ref_blocks);
            lista_butoane.Add(button_generate_excel);
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
            lista_butoane.Add(button_replace_block);
            lista_butoane.Add(button_load_reference_library);
            lista_butoane.Add(button_select_ref_blocks);
            lista_butoane.Add(button_generate_excel);

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

            var toolStripMenuItem5 = new ToolStripMenuItem { Text = "Add Column A dwgs to drawing list" };
            toolStripMenuItem5.Click += if_there_is_a_path_add_to_dt_display_click;

            var toolStripMenuItem6 = new ToolStripMenuItem { Text = "Open all drawings" };
            toolStripMenuItem6.Click += open_all_drawings_Click;


            ContextMenuStrip_open_alignment = new ContextMenuStrip();
            ContextMenuStrip_open_alignment.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2, toolStripMenuItem3, toolStripMenuItem4, toolStripMenuItem5, toolStripMenuItem6 });


            var toolStripMenuItem7 = new ToolStripMenuItem { Text = "Select" };
            toolStripMenuItem7.Click += Select_cell_Click;

            var toolStripMenuItem8 = new ToolStripMenuItem { Text = "Unselect" };
            toolStripMenuItem8.Click += Unselect_cell_Click;


            var toolStripMenuItem9 = new ToolStripMenuItem { Text = "Unselect All" };
            toolStripMenuItem9.Click += Unselect_all_cells_Click;



            ContextMenuStrip_references = new ContextMenuStrip();
            ContextMenuStrip_references.Items.AddRange(new ToolStripItem[] { toolStripMenuItem7, toolStripMenuItem8, toolStripMenuItem9 });

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

                            string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();
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

        private void open_all_drawings_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_drawings.RowCount > 0)
                {

                    for (int i = 0; i < Display_dataTable.Rows.Count; i++)
                    {
                        if (Display_dataTable != null)
                        {
                            if (Display_dataTable.Rows[i][0] != DBNull.Value)
                            {
                                string fisier_generat = Convert.ToString(Display_dataTable.Rows[i][0]);

                                string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();
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
            try
            {
                set_enable_false();
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _AGEN_mainform.block_attributes_excel_name;

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

                    Microsoft.Office.Interop.Excel.Application Excel1 = null;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;


                    if (System.IO.File.Exists(file_de_procesat) == true)
                    {

                        Functions.create_backup(file_de_procesat);


                        try
                        {
                            Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        }
                        catch (System.Exception ex)
                        {
                            Excel1 = new Microsoft.Office.Interop.Excel.Application();

                        }

                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                        Workbook1 = Excel1.Workbooks.Open(file_de_procesat);
                        W1 = Workbook1.Worksheets[1];

                        try
                        {


                            System.Data.DataTable dt_ex = Load_attributes_from_excel(W1, null, null, false);

                            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);
                            System.Data.DataTable dt_new = creaza_data_table_attributes(W1);
                            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt2);

                            if (dt_ex != null)
                            {
                                List<string> lista1 = new List<string>();
                                lista1.Add("Drawing");
                                lista1.Add("Layout");

                                for (int j = 2; j < dt_ex.Columns.Count; j++)
                                {
                                    string block_name1 = Convert.ToString(dt_ex.Rows[0][j]);
                                    string atrib_name1 = Convert.ToString(dt_ex.Rows[1][j]);
                                    string colname1 = block_name1 + "|" + atrib_name1;
                                    lista1.Add(colname1);
                                }

                                for (int i = 2; i < dt_ex.Rows.Count; i++)
                                {
                                    dt_new.Rows.Add();
                                    for (int j = 0; j < 2; j++)
                                    {
                                        dt_new.Rows[dt_new.Rows.Count - 1][j] = dt_ex.Rows[i][j];
                                    }
                                    for (int k = 0; k < lista1.Count; k++)
                                    {
                                        if (dt_new.Columns.Contains(lista1[k]) == true)
                                        {
                                            for (int j = 2; j < dt_ex.Columns.Count; j++)
                                            {
                                                string block_name1 = Convert.ToString(dt_ex.Rows[0][j]);
                                                string atrib_name1 = Convert.ToString(dt_ex.Rows[1][j]);
                                                string colname1 = block_name1 + "|" + atrib_name1;
                                                if (lista1[k] == colname1)
                                                {
                                                    dt_new.Rows[dt_new.Rows.Count - 1][lista1[k]] = dt_ex.Rows[i][j];
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (dt_new != null)
                            {
                                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                if (segment1 == "not defined") segment1 = "";
                                Functions.Create_header_block_attributes_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1, dt_new.Columns.Count);
                                System.Data.DataTable dt_header = creaza_data_table_excel_header_values(dt_new);
                                Transfera_data_to_excel_fara_header(W1, dt_header, _AGEN_mainform.Start_row_block_attributes);
                                Transfer_to_w1_Data_table_values(W1, dt_new, _AGEN_mainform.Start_row_block_attributes + 2);
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

            object[,] matrix_w1_header = new object[2, 648];
            matrix_w1_header = range1.Value2;

            dt1.Columns.Add("Drawing", typeof(string));
            dt1.Columns.Add("Layout", typeof(string));

            Display_dataTable = new System.Data.DataTable();
            Display_dataTable.Columns.Add("Drawing", typeof(string));
            Display_dataTable.Columns.Add("Layout", typeof(string));
            Display_dataTable.Columns.Add("Blocks Found", typeof(string));

            int end1 = 0;
            for (int i = 1; i <= 648; ++i)
            {
                object val1 = matrix_w1_header[1, i];
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

            int duplicate_idx = 1;
            for (int i = 3; i <= end1; ++i)
            {
                string colname = matrix_w1_header[1, i].ToString() + "|" + matrix_w1_header[2, i].ToString();
                if (dt1.Columns.Contains(colname) == true)
                {
                    do
                    {
                        colname = colname + "_d" + Convert.ToString(duplicate_idx);
                        ++duplicate_idx;
                    } while (dt1.Columns.Contains(colname) == true);
                }
                dt1.Columns.Add(colname, typeof(string));
                duplicate_idx = 1;
            }

            string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();

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
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        drawing_list = new List<string>();
                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                drawing_list.Add(file1);
                            }
                        }
                    }
                }

                if (drawing_list.Count == 0)
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
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

                                                                if (dt1.Rows[dt1.Rows.Count - 1][colname] != DBNull.Value)
                                                                {
                                                                    string colname1 = colname + "_d" + Convert.ToString(duplicate_idx);
                                                                    if (dt1.Columns.Contains(colname1) == true)
                                                                    {
                                                                        do
                                                                        {
                                                                            if (dt1.Rows[dt1.Rows.Count - 1][colname1] != DBNull.Value)
                                                                            {
                                                                                ++duplicate_idx;
                                                                                colname1 = colname1 + "_d" + Convert.ToString(duplicate_idx);
                                                                            }
                                                                            else
                                                                            {
                                                                                duplicate_idx = 1;
                                                                            }

                                                                        } while (dt1.Columns.Contains(colname1) == true);
                                                                        colname = colname1;
                                                                    }
                                                                }

                                                                string[] bl_atr = colname.Split(split1);


                                                                if (block_name.ToLower() == bl_atr[0].ToLower())
                                                                {
                                                                    found1 = "YES";

                                                                    if (bl_atr[1].Contains("_d") == true)
                                                                    {
                                                                        int index1 = bl_atr[1].IndexOf("_d");
                                                                        bl_atr[1] = bl_atr[1].Substring(0, index1 + 1);
                                                                    }

                                                                    foreach (ObjectId atid in block1.AttributeCollection)
                                                                    {
                                                                        AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForRead);

                                                                        if (Atr1 != null)
                                                                        {
                                                                            if (Atr1.Tag.ToLower() == bl_atr[1].ToLower() && Atr1.TextString != "")
                                                                            {
                                                                                dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.TextString;
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

                                        if (checkBox_ms.Checked == true)
                                        {

                                            dt1.Rows.Add();
                                            Display_dataTable.Rows.Add();
                                           string found1 = "NO";
                                            BlockTable BlockTable1 = Database2.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                            BlockTableRecord BtrecordMS = (BlockTableRecord)Trans2.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                                            foreach (ObjectId id1 in BtrecordMS)
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

                                                                if (bl_atr[1].Contains("_d") == true)
                                                                {
                                                                    int index1 = bl_atr[1].IndexOf("_d");
                                                                    bl_atr[1] = bl_atr[1].Substring(0, index1 + 1);
                                                                }

                                                                foreach (ObjectId atid in block1.AttributeCollection)
                                                                {
                                                                    AttributeReference Atr1 = (AttributeReference)Trans2.GetObject(atid, OpenMode.ForRead);

                                                                    if (Atr1 != null)
                                                                    {
                                                                        if (Atr1.Tag.ToLower() == bl_atr[1].ToLower() && Atr1.TextString != "")
                                                                        {
                                                                            dt1.Rows[dt1.Rows.Count - 1][colname] = Atr1.TextString;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            dt1.Rows[dt1.Rows.Count - 1]["Drawing"] = file1;
                                            dt1.Rows[dt1.Rows.Count - 1]["Layout"] = "MODEL";

                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Drawing"] = file1;
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Layout"] = "MODEL";
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Blocks Found"] = found1;
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
            Functions.Kill_excel();

            Display_dataTable = new System.Data.DataTable();
            Display_dataTable.Columns.Add("Drawing", typeof(string));
            Display_dataTable.Columns.Add("Layout", typeof(string));
            Display_dataTable.Columns.Add("Blocks Populated", typeof(string));




            try
            {
                set_enable_false();
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _AGEN_mainform.block_attributes_excel_name;

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
                    Microsoft.Office.Interop.Excel.Worksheet[] W2 = null;
                    Microsoft.Office.Interop.Excel.Worksheet[] W3 = null;
                    int no_w2 = 0;
                    int no_w3 = 0;
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

                                if (Workbook1.Worksheets.Count > 1)
                                {
                                    for (int i = 2; i <= Workbook1.Worksheets.Count; ++i)
                                    {
                                        Microsoft.Office.Interop.Excel.Worksheet W11 = Workbook1.Worksheets[i];

                                        if (W11.Name.ToUpper().Contains("VER") == true)
                                        {
                                            ++no_w2;
                                            Array.Resize(ref W2, no_w2);
                                            W2[no_w2 - 1] = W11;
                                        }
                                        if (W11.Name.ToUpper().Contains("HOR") == true)
                                        {
                                            ++no_w3;
                                            Array.Resize(ref W3, no_w3);
                                            W3[no_w3 - 1] = W11;
                                        }
                                    }
                                }



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

                            if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                            Workbook1 = Excel1.Workbooks.Open(file_de_procesat);
                            W1 = Workbook1.Worksheets[1];
                            if (Workbook1.Worksheets.Count > 1)
                            {
                                for (int i = 2; i <= Workbook1.Worksheets.Count - 1; ++i)
                                {
                                    Microsoft.Office.Interop.Excel.Worksheet W11 = Workbook1.Worksheets[i];

                                    if (W11.Name.ToUpper().Contains("VER") == true)
                                    {
                                        ++no_w2;
                                        Array.Resize(ref W2, no_w2);
                                        W2[no_w2 - 1] = W11;
                                    }
                                    if (W11.Name.ToUpper().Contains("HOR") == true)
                                    {
                                        ++no_w3;
                                        Array.Resize(ref W3, no_w3);
                                        W3[no_w3 - 1] = W11;
                                    }
                                }
                            }
                        }



                        try
                        {

                            System.Data.DataTable dt1 = Load_attributes_from_excel(W1, W2, W3);

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

                                                        string path1 = opened_dwg.Database.OriginalFileName;




                                                        if (path1 == file1)
                                                        {
                                                            HostApplicationServices.WorkingDatabase = opened_dwg.Database;
                                                            document_collection.MdiActiveDocument = opened_dwg;
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

                                                            if (checkBox_ac1024.Checked == false)
                                                            {
                                                                Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                            }
                                                            else
                                                            {
                                                                Database2.SaveAs(file1, true, DwgVersion.AC1024, ThisDrawing.Database.SecurityParameters);
                                                            }
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
                            if (block1.AttributeCollection.Count > 0)
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
                        if (block1.AttributeCollection.Count > 0)
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
                                }
                            }
                        }
                    }
                }

                Trans2.Commit();
                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1]["Blocks Populated"] = found1;
            }
        }

        private System.Data.DataTable Load_attributes_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, Microsoft.Office.Interop.Excel.Worksheet[] W22 = null, Microsoft.Office.Interop.Excel.Worksheet[] W33 = null, bool show_messagebox = true)
        {


            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Drawing", typeof(string));
            dt1.Columns.Add("Layout", typeof(string));



            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A7:XX8"];
            object[,] matrix_W1_header = new object[2, 648];
            matrix_W1_header = range1.Value2;



            int col_end1 = 2;
            for (int j = 3; j <= 648; ++j)
            {
                object val1 = matrix_W1_header[1, j];
                object val2 = matrix_W1_header[2, j];

                if (val1 == null || val2 == null)
                {
                    col_end1 = j - 1;
                    j = 649;
                }
            }

            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A1:A30000"];

            object[,] matrix_w1_rows = new object[1, 30000];
            matrix_w1_rows = range2.Value2;



            int row_end1 = 8;
            for (int i = 9; i <= 30000; ++i)
            {
                object val1 = matrix_w1_rows[i, 1];

                if (val1 == null)
                {
                    row_end1 = i - 1;
                    i = 30001;
                }
            }

            if (col_end1 < 3)
            {
                if (show_messagebox == true) MessageBox.Show("no header found for the block attributes file\r\nOperation aborted");
                return null;
            }

            if (row_end1 < 9)
            {
                if (show_messagebox == true) MessageBox.Show("no attribute data found\r\nOperation aborted");
                return null;
            }

            for (int j = 3; j <= col_end1; ++j)
            {
                dt1.Columns.Add("col_" + j.ToString(), typeof(string));
            }


            Microsoft.Office.Interop.Excel.Range range3 = W1.Range[W1.Cells[1, 1], W1.Cells[row_end1, col_end1]];

            object[,] matrix_w1_all_data = new object[row_end1, col_end1];
            matrix_w1_all_data = range3.Value2;

            for (int i = 7; i <= 8; ++i)
            {
                dt1.Rows.Add();

                dt1.Rows[dt1.Rows.Count - 1]["Drawing"] = "***";
                dt1.Rows[dt1.Rows.Count - 1]["Layout"] = "***";

                for (int j = 3; j <= col_end1; ++j)
                {
                    object val1 = matrix_w1_all_data[i, j];
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
                    object val1 = matrix_w1_all_data[i, j];
                    if (val1 != null)
                    {
                        dt1.Rows[dt1.Rows.Count - 1][j - 1] = val1.ToString();
                    }
                }
            }



            if (W33 != null)
            {
                for (int k = 0; k < W33.Length; k++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W3 = W33[k];
                    Microsoft.Office.Interop.Excel.Range range31 = W3.Range["A1:XX2"];
                    object[,] matrix_w3_header = new object[2, 648];
                    matrix_w3_header = range31.Value2;
                    int col_end3 = 2;
                    for (int j = 3; j <= 648; ++j)
                    {
                        object val1 = matrix_w3_header[1, j];
                        object val2 = matrix_w3_header[2, j];
                        if (val1 == null || val2 == null)
                        {
                            col_end3 = j - 1;
                            j = 649;
                        }
                    }

                    if (col_end3 > 2)
                    {
                        for (int j = 3; j <= col_end3; ++j)
                        {
                            dt1.Columns.Add("colH_" + j.ToString() + k.ToString(), typeof(string));
                        }
                        Microsoft.Office.Interop.Excel.Range range33 = W3.Range[W3.Cells[1, 1], W3.Cells[row_end1 - 6, col_end3]];
                        object[,] matrix_w3 = new object[row_end1 - 6, col_end3];
                        matrix_w3 = range33.Value2;
                        for (int i = 1; i <= 2; ++i)
                        {
                            for (int j = 3; j <= col_end3; ++j)
                            {
                                object val1 = matrix_w3[i, j];
                                if (val1 != null)
                                {
                                    dt1.Rows[i - 1]["colH_" + j.ToString() + k.ToString()] = val1.ToString();
                                }
                            }
                        }
                        for (int i = 3; i <= row_end1 - 6; ++i)
                        {
                            for (int j = 3; j <= col_end3; ++j)
                            {
                                object val1 = matrix_w3[i, j];
                                if (val1 != null)
                                {
                                    dt1.Rows[i - 1]["colH_" + j.ToString() + k.ToString()] = val1.ToString();
                                }
                            }
                        }
                    }
                }
            }
            //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1);

            if (W22 != null)
            {
                for (int k = 0; k < W22.Length; k++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W2 = W22[k];
                    Microsoft.Office.Interop.Excel.Range range21 = W2.Range["A1:B30000"];
                    object[,] matrix_w2_header = new object[30000, 2];
                    matrix_w2_header = range21.Value2;
                    int col_end2 = 1;
                    for (int i = 2; i <= 30000; ++i)
                    {
                        object val1 = matrix_w2_header[i, 1];
                        object val2 = matrix_w2_header[i, 2];
                        if (val1 == null || val2 == null)
                        {
                            col_end2 = i - 1;
                            i = 30649;
                        }
                    }

                    if (col_end2 > 1)
                    {
                        for (int i = 2; i <= col_end2; ++i)
                        {
                            dt1.Columns.Add("colV_" + i.ToString() + k.ToString(), typeof(string));
                        }
                        Microsoft.Office.Interop.Excel.Range range22 = W2.Range[W2.Cells[1, 1], W2.Cells[col_end2, row_end1 - 6]];
                        object[,] matrix_w2 = new object[col_end2, row_end1 - 6];
                        matrix_w2 = range22.Value2;

                        for (int i = 2; i <= col_end2; ++i)
                        {
                            for (int j = 1; j <= row_end1 - 6; ++j)
                            {
                                object val1 = matrix_w2[i, j];
                                if (val1 != null)
                                {
                                    dt1.Rows[j - 1]["colV_" + i.ToString() + k.ToString()] = val1.ToString();
                                }
                            }
                        }

                    }
                }
            }


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
                    if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                    {
                        string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }
                        string fisier_attributes = ProjF + _AGEN_mainform.block_attributes_excel_name;

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

                    _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                    _AGEN_mainform.tpage_processing.Show();

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
                    _AGEN_mainform.tpage_processing.Hide();
                    Ag.WindowState = FormWindowState.Normal;
                }
            }
        }

        private void button_load_block_attributes_to_excel_header_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();


            if (Functions.Get_if_workbook_is_open_in_Excel(Excel_tblk_atr) == true)
            {
                MessageBox.Show("Please close the " + Excel_tblk_atr + " file");
                return;
            }

            Ag = this.MdiParent as _AGEN_mainform;

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
                                if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                                {
                                    string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                                    {
                                        ProjF = ProjF + "\\";
                                    }
                                    string fisier_atr = ProjF + _AGEN_mainform.block_attributes_excel_name;

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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
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
                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                    if (segment1 == "not defined") segment1 = "";
                    Functions.Create_header_block_attributes_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1, dt_atr.Columns.Count);
                    Transfera_data_to_excel_fara_header(W1, dt_atr, _AGEN_mainform.Start_row_block_attributes);

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
            Functions.Kill_excel();


            if (Functions.Get_if_workbook_is_open_in_Excel(Excel_tblk_atr) == true)
            {
                MessageBox.Show("Please close the " + Excel_tblk_atr + " file");
                return;
            }

            try
            {
                set_enable_false();
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _AGEN_mainform.block_attributes_excel_name;

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

                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
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

            string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();

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
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                    {
                        drawing_list = new List<string>();
                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                            {
                                string file1 = _AGEN_mainform.dt_sheet_index.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
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




        private void if_there_is_a_path_add_to_dt_display_click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                if (Ag != null)
                {

                    string file_de_procesat = "";

                    if (Excel_tblk_atr == "")
                    {
                        if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                        {
                            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                            {
                                ProjF = ProjF + "\\";
                            }

                            string fisier_attributes = ProjF + _AGEN_mainform.block_attributes_excel_name;

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

                            if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
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



                            if (Display_dataTable == null)
                            {
                                Display_dataTable = new System.Data.DataTable();
                                Display_dataTable.Columns.Add("Drawing", typeof(string));
                            }

                            if (drawing_list == null)
                            {
                                drawing_list = new List<string>();
                            }


                            for (int i = 0; i < dt1.Rows.Count; i++)
                            {
                                if (dt1.Rows[i]["Drawing"] != DBNull.Value)
                                {
                                    string dwg1 = Convert.ToString(dt1.Rows[i]["Drawing"]);
                                    if (System.IO.File.Exists(dwg1) == true)
                                    {
                                        if (drawing_list.Contains(dwg1) == false)
                                        {
                                            Display_dataTable.Rows.Add();
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][0] = dwg1;
                                            drawing_list.Add(dwg1);
                                        }
                                    }
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

                            _AGEN_mainform.tpage_processing.Hide();
                            Ag.WindowState = FormWindowState.Normal;


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
            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.block_attributes_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.block_attributes_excel_name + " file");
                return;
            }
            string fisier_block_atr = "";
            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();
                if (ProjFolder.Length > 0 && ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    fisier_block_atr = ProjFolder + _AGEN_mainform.block_attributes_excel_name;
                    if (System.IO.File.Exists(fisier_block_atr) == true)
                    {
                        MessageBox.Show("there is already a " + _AGEN_mainform.block_attributes_excel_name + " excel file");
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
                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;


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
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
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
                    MessageBox.Show("there is already a " + _AGEN_mainform.block_attributes_excel_name + " excel file");
                    return;
                }
            }
        }

        private void button_replace_block_Click(object sender, EventArgs e)
        {




            if (Display_dataTable != null && Display_dataTable.Rows.Count > 0)
            {
                drawing_list = new List<string>();
                for (int i = 0; i < Display_dataTable.Rows.Count; ++i)
                {
                    drawing_list.Add(Convert.ToString(Display_dataTable.Rows[i][0]));
                }




                Display_dataTable = new System.Data.DataTable();
                Display_dataTable.Columns.Add("Drawing", typeof(string));
                Display_dataTable.Columns.Add("Block Deleted", typeof(string));
                Display_dataTable.Columns.Add("Blocks Replaced", typeof(string));
                Display_dataTable.Columns.Add("Insert New Block", typeof(string));





                try
                {
                    set_enable_false();
                    _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                    if (Ag != null)
                    {

                        string new_file_block = textBox_block_name.Text;





                        if (System.IO.File.Exists(new_file_block) == true)
                        {





                            try
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


                                    if (drawing_list != null && drawing_list.Count > 0)
                                    {
                                        for (int i = 0; i < drawing_list.Count; ++i)
                                        {



                                            string file1 = drawing_list[i];





                                            if (System.IO.File.Exists(file1) == true)
                                            {
                                                Display_dataTable.Rows.Add();
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][0] = file1;
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][1] = "NO";
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][2] = "NO";
                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][3] = "NO";

                                                bool is_opened = false;
                                                DocumentCollection document_collection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                                                foreach (Document opened_dwg in document_collection)
                                                {
                                                    if (opened_dwg.Database.Filename == file1)
                                                    {

                                                        is_opened = true;
                                                        document_collection.MdiActiveDocument = opened_dwg;
                                                        HostApplicationServices.WorkingDatabase = opened_dwg.Database;
                                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans3 = opened_dwg.TransactionManager.StartTransaction())
                                                        {
                                                            bool isfound = false;
                                                            bool isreplaced = false;



                                                            replace_block(new_file_block, Trans3, opened_dwg.Database, ref isfound, ref isreplaced);
                                                            if (isfound == true)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][1] = "YES";
                                                            }
                                                            if (isreplaced == true)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][2] = "YES";
                                                            }

                                                            if (isfound == false && isreplaced == false)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][3] = "YES";

                                                            }

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
                                                            bool isfound = false;
                                                            bool isreplaced = false;
                                                            replace_block_closed_file(new_file_block, Trans2, Database2, ref isfound, ref isreplaced);
                                                            if (isfound == true)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][1] = "YES";
                                                            }
                                                            if (isreplaced == true)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][2] = "YES";
                                                            }

                                                            if (isfound == false && isreplaced == false)
                                                            {
                                                                Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][3] = "YES";

                                                            }

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
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.Message);

                            }
                        }
                        else
                        {
                            dataGridView_drawings.DataSource = null;
                            Display_dataTable = new System.Data.DataTable();
                            drawing_list = new List<string>();
                            MessageBox.Show("you do not have the new block file", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        }
        public void replace_block(string path1, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, ref bool isfound, ref bool isreplaced)
        {

            string new_block_name = System.IO.Path.GetFileNameWithoutExtension(path1);

            try
            {



                using (Trans1)
                {
                    BlockTable BlockTable1 = Database1.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                    BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    System.Data.DataTable dt1 = new System.Data.DataTable();

                    dt1.Columns.Add("LAYER", typeof(string));
                    dt1.Columns.Add("IP", typeof(Point3d));
                    dt1.Columns.Add("SCALE", typeof(double));
                    dt1.Columns.Add("ROT", typeof(double));
                    dt1.Columns.Add("CI", typeof(int));



                    foreach (ObjectId id1 in BTrecord_PS)
                    {
                        BlockReference bl1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;

                        if (bl1 != null)
                        {
                            string old1 = Functions.get_block_name(bl1);
                            if (old1.ToLower() == new_block_name.ToLower())
                            {



                                string layer1 = bl1.Layer;
                                Point3d inspt = bl1.Position;
                                double rot = bl1.Rotation;
                                int ci = bl1.ColorIndex;
                                double xscale = bl1.ScaleFactors.X;

                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1]["LAYER"] = layer1;
                                dt1.Rows[dt1.Rows.Count - 1]["IP"] = inspt;
                                dt1.Rows[dt1.Rows.Count - 1]["SCALE"] = xscale;
                                dt1.Rows[dt1.Rows.Count - 1]["ROT"] = rot;
                                dt1.Rows[dt1.Rows.Count - 1]["CI"] = ci;

                                bl1.UpgradeOpen();
                                bl1.Erase();
                                isfound = true;
                            }
                        }

                    }


                    erase_block_definition_from_database(Database1, new_block_name);


                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {

                        string layer1 = Convert.ToString(dt1.Rows[i]["LAYER"]);
                        double rot = Convert.ToDouble(dt1.Rows[i]["ROT"]);
                        double xscale = Convert.ToDouble(dt1.Rows[i]["SCALE"]);
                        Point3d inspt = (Point3d)dt1.Rows[dt1.Rows.Count - 1]["IP"];
                        int ci = Convert.ToInt32(dt1.Rows[i]["CI"]);


                        BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(Database1, BTrecord_PS, path1, new_block_name, inspt, xscale, rot, layer1, col_atr, col_val);
                        block1.ColorIndex = ci;
                        isreplaced = true;
                    }

                    if (dt1.Rows.Count == 0)
                    {
                        Point3d inspt = new Point3d(0, 0, 0);
                        BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(Database1, BTrecord_PS, path1, new_block_name, inspt, 1, 0, "0", col_atr, col_val);
                        block1.ColorIndex = 256;
                    }



                    Trans1.Commit();
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }



        public void replace_block_closed_file(string path1, Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, Database Database1, ref bool isfound, ref bool isreplaced)
        {

            string new_block_name = System.IO.Path.GetFileNameWithoutExtension(path1);

            try
            {



                using (Trans1)
                {
                    BlockTable BlockTable1 = Database1.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                    Functions.make_first_layout_active(Trans1, Database1);
                    BlockTableRecord BTrecord_PS = Functions.get_first_layout_as_paperspace(Trans1, Database1);
                    BTrecord_PS.UpgradeOpen();


                    System.Data.DataTable dt1 = new System.Data.DataTable();

                    dt1.Columns.Add("LAYER", typeof(string));
                    dt1.Columns.Add("IP", typeof(Point3d));
                    dt1.Columns.Add("SCALE", typeof(double));
                    dt1.Columns.Add("ROT", typeof(double));
                    dt1.Columns.Add("CI", typeof(int));



                    foreach (ObjectId id1 in BTrecord_PS)
                    {
                        BlockReference bl1 = Trans1.GetObject(id1, OpenMode.ForRead) as BlockReference;

                        if (bl1 != null)
                        {
                            string old1 = Functions.get_block_name_another_database(bl1, Database1);
                            if (old1.ToLower() == new_block_name.ToLower())
                            {



                                string layer1 = bl1.Layer;
                                Point3d inspt = bl1.Position;
                                double rot = bl1.Rotation;
                                int ci = bl1.ColorIndex;
                                double xscale = bl1.ScaleFactors.X;

                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1]["LAYER"] = layer1;
                                dt1.Rows[dt1.Rows.Count - 1]["IP"] = inspt;
                                dt1.Rows[dt1.Rows.Count - 1]["SCALE"] = xscale;
                                dt1.Rows[dt1.Rows.Count - 1]["ROT"] = rot;
                                dt1.Rows[dt1.Rows.Count - 1]["CI"] = ci;

                                bl1.UpgradeOpen();
                                bl1.Erase();
                                isfound = true;
                            }
                        }

                    }


                    erase_block_definition_from_database(Database1, new_block_name);


                    System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                    System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();
                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {

                        string layer1 = Convert.ToString(dt1.Rows[i]["LAYER"]);
                        double rot = Convert.ToDouble(dt1.Rows[i]["ROT"]);
                        double xscale = Convert.ToDouble(dt1.Rows[i]["SCALE"]);
                        Point3d inspt = (Point3d)dt1.Rows[dt1.Rows.Count - 1]["IP"];
                        int ci = Convert.ToInt32(dt1.Rows[i]["CI"]);


                        BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(Database1, BTrecord_PS, path1, new_block_name, inspt, xscale, rot, layer1, col_atr, col_val);
                        block1.ColorIndex = ci;
                        isreplaced = true;
                    }

                    if (dt1.Rows.Count == 0)
                    {
                        Point3d inspt = new Point3d(0, 0, 0);
                        BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(Database1, BTrecord_PS, path1, new_block_name, inspt, 1, 0, "0", col_atr, col_val);
                        block1.ColorIndex = 256;
                    }



                    Trans1.Commit();
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }


        private void erase_block_definition_from_database(Database Database1, string NumeBlock)
        {
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = Database1.TransactionManager.StartTransaction())
            {

                BlockTable BlockTable1 = (BlockTable)Trans1.GetObject(Database1.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                if (BlockTable1.Has(NumeBlock) == true)
                {
                    BlockTableRecord btr1 = Trans1.GetObject(BlockTable1[NumeBlock], OpenMode.ForWrite) as BlockTableRecord;
                    if (btr1 != null)
                    {
                        btr1.Erase();
                    }


                }
                Trans1.Commit();
            }


        }

        private void label_station_equations_Click(object sender, EventArgs e)
        {
            if (label_file_name.Visible == false)
            {
                label_file_name.Visible = true;
                button_replace_block.Visible = true;
                textBox_block_name.Visible = true;
                button_dan.Visible = true;

            }
            else
            {
                label_file_name.Visible = false;
                button_replace_block.Visible = false;
                textBox_block_name.Visible = false;
                button_dan.Visible = false;


            }

        }

        private void panel_dan_Click(object sender, EventArgs e)
        {
            if (Functions.is_dan_popescu() == true)
            {
                if (panel_dan.Visible == false)
                {
                    panel_dan.Visible = true;
                    panel_ref.Visible = true;
                }
                else
                {
                    panel_dan.Visible = false;
                    panel_ref.Visible = false;

                }
            }
        }

        private void button_load_reference_library_Click(object sender, EventArgs e)
        {

            _AGEN_mainform.dt_ref = get_dt_ref_structure();
            string wn = "AGEN_reference.xlsx";
            bool excel_file_opened = false;
            set_enable_false();
            try
            {

                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                try
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();

                    }

                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;

                    if (Excel1 != null)
                    {

                        foreach (Workbook wkb1 in Excel1.Workbooks)
                        {

                            if (wkb1.Name == wn)
                            {
                                Workbook1 = wkb1;
                                W1 = wkb1.Worksheets[1];
                                excel_file_opened = true;

                            }

                        }
                    }

                    if (W1 == null)
                    {
                        if (System.IO.File.Exists(_AGEN_mainform.ProjFolder + wn) == true)
                        {
                            Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.ProjFolder + wn);
                            W1 = Workbook1.Worksheets[1];
                        }
                    }
                    if (W1 != null)
                    {
                        Range range1 = W1.Range["A1:E30000"];
                        object[,] values1 = new object[30000, 5];
                        values1 = range1.Value2;

                        for (int i = 2; i <= 30000; ++i)
                        {
                            object val1 = values1[i, 1];
                            object val2 = values1[i, 2];
                            object val3 = values1[i, 3];
                            object val4 = values1[i, 4];
                            object val5 = values1[i, 5];

                            if (val2 != null)
                            {
                                if (val1 != null)
                                {
                                    _AGEN_mainform.dt_ref.Rows.Add();
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][0] = Convert.ToString(val1);
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][1] = Convert.ToString(val2);
                                }
                                else
                                {
                                    string line1 = Convert.ToString(_AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][1]);
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][1] = line1 + '|' + Convert.ToString(val2);
                                }
                                if (val3 != null && Functions.IsNumeric(val3.ToString()) == true)
                                {
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][2] = Convert.ToDouble(val3);
                                }
                                if (val4 != null && Functions.IsNumeric(val4.ToString()) == true)
                                {
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][3] = Convert.ToDouble(val4);
                                }
                                if (val5 != null)
                                {
                                    _AGEN_mainform.dt_ref.Rows[_AGEN_mainform.dt_ref.Rows.Count - 1][4] = Convert.ToString(val5);
                                }
                            }
                            else
                            {
                                i = values1.Length + 1;
                            }
                        }

                        if (excel_file_opened == false)
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
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("no excel found");

                }
                finally
                {
                    if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                    if (excel_file_opened == false) if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (excel_file_opened == false) if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }

            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            set_enable_true();

            _AGEN_mainform.dt_ref = Functions.Sort_data_table(_AGEN_mainform.dt_ref, "Dwg");

            dataGridView_ref.DataSource = _AGEN_mainform.dt_ref;
            dataGridView_ref.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_ref.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_ref.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_ref.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_ref.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_ref.EnableHeadersVisualStyles = false;

        }

        private System.Data.DataTable get_dt_ref_structure()
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Dwg", typeof(string));
            dt1.Columns.Add("Description", typeof(string));
            dt1.Columns.Add("Start", typeof(double));
            dt1.Columns.Add("End", typeof(double));
            dt1.Columns.Add("Block Name", typeof(string));

            return dt1;
        }


        private void button_select_ref_blocks_Click(object sender, EventArgs e)
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
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the blocks:";
                        Prompt_rez.SingleOnly = false;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status != PromptStatus.OK)
                        {

                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }
                        double sta_min = -1.234;
                        double sta_max = -1.234;
                        List<string> lista1 = new List<string>();
                        List<string> lista2 = new List<string>();
                        for (int i = 0; i < Rezultat1.Value.Count; i++)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;

                            if (block1 != null)
                            {
                                string nume2 = Functions.get_block_name(block1);
                                if (lista2.Contains(nume2) == false)
                                {
                                    lista2.Add(nume2);
                                }


                                if (block1.AttributeCollection.Count > 0)
                                {



                                    foreach (ObjectId id1 in block1.AttributeCollection)
                                    {
                                        AttributeReference atr1 = Trans1.GetObject(id1, OpenMode.ForRead) as AttributeReference;
                                        if (atr1.Tag.ToUpper() == "ID")
                                        {
                                            if (atr1.TextString != "")
                                            {
                                                if (lista1.Contains(atr1.TextString) == false)
                                                {
                                                    lista1.Add(atr1.TextString);
                                                }
                                            }
                                        }

                                        if (atr1.Tag.ToUpper() == "STA" || atr1.Tag.ToUpper() == "STA1" || atr1.Tag.ToUpper() == "STA2" || atr1.Tag.ToUpper() == "STA11" || atr1.Tag.ToUpper() == "STA21")
                                        {
                                            if (Functions.IsNumeric(atr1.TextString.Replace("+", "")) == true)
                                            {
                                                double sta = Convert.ToDouble(atr1.TextString.Replace("+", ""));
                                                if (sta_min == -1.234)
                                                {
                                                    sta_min = sta;
                                                }
                                                if (sta_max == -1.234)
                                                {
                                                    sta_max = sta;
                                                }

                                                if (sta > sta_max) sta_max = sta;
                                                if (sta < sta_min) sta_min = sta;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (lista1.Count > 0)
                        {
                            for (int i = 0; i < lista1.Count; i++)
                            {
                                string ref1 = lista1[i];

                                for (int j = 0; j < dataGridView_ref.Rows.Count; j++)
                                {
                                    if (dataGridView_ref.Rows[j].Cells[0].Value != DBNull.Value)
                                    {
                                        string ref2 = Convert.ToString(dataGridView_ref.Rows[j].Cells[0].Value);
                                        if (ref1 == ref2)
                                        {
                                            dataGridView_ref.Rows[j].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Bold);
                                            dataGridView_ref.Rows[j].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                            dataGridView_ref.Rows[j].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                                        }
                                        else
                                        {
                                            double Sta1 = -1.234;
                                            double Sta2 = -1.234;
                                            if (_AGEN_mainform.dt_ref.Rows[j][2] != DBNull.Value)
                                            {
                                                Sta1 = Convert.ToDouble(_AGEN_mainform.dt_ref.Rows[j][2]);
                                            }
                                            if (_AGEN_mainform.dt_ref.Rows[j][3] != DBNull.Value)
                                            {
                                                Sta2 = Convert.ToDouble(_AGEN_mainform.dt_ref.Rows[j][3]);
                                            }
                                            bool select_cell = false;
                                            if (Sta1 != -1.234 && Sta2 != -1.234)
                                            {
                                                if (sta_min < Sta2 && sta_min >= Sta1)
                                                {
                                                    select_cell = true;
                                                }

                                                if (sta_max <= Sta2 && sta_max > Sta1)
                                                {
                                                    select_cell = true;
                                                }

                                                if (sta_min > Sta1 && sta_max < Sta2)
                                                {
                                                    select_cell = true;
                                                }

                                                if (sta_min < Sta1 && sta_max > Sta2)
                                                {
                                                    select_cell = true;
                                                }
                                            }

                                            if (_AGEN_mainform.dt_ref.Rows[j][4] != DBNull.Value)
                                            {
                                                string nume1 = Convert.ToString(_AGEN_mainform.dt_ref.Rows[j][4]);
                                                if (lista2.Contains(nume1) == true)
                                                {
                                                    select_cell = true;
                                                }
                                            }

                                            if (select_cell == true)
                                            {
                                                dataGridView_ref.Rows[j].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Bold);
                                                dataGridView_ref.Rows[j].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                                                dataGridView_ref.Rows[j].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                                            }
                                        }
                                    }
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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }


        private void dataGridView_ref_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_ref.CurrentCell = dataGridView_ref.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_references.Show(Cursor.Position);
                ContextMenuStrip_references.Visible = true;
            }
            else
            {
                ContextMenuStrip_references.Visible = false;
            }
            if (e.Button == System.Windows.Forms.MouseButtons.Left && e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                dataGridView_ref.CurrentCell = dataGridView_ref.Rows[e.RowIndex].Cells[e.ColumnIndex];

                if (dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.ForeColor == Color.FromArgb(0, 0, 0))
                {
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Regular);
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.White;
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.BackColor = Color.FromArgb(37, 37, 38);
                }
                else
                {
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Bold);
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                    dataGridView_ref.Rows[e.RowIndex].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                }
            }

        }

        private void Select_cell_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_ref.RowCount > 0)
                {
                    int idx1 = dataGridView_ref.CurrentCell.RowIndex;
                    if (idx1 >= 0)
                    {
                        dataGridView_ref.Rows[idx1].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Bold);
                        dataGridView_ref.Rows[idx1].Cells[0].Style.ForeColor = Color.FromArgb(0, 0, 0);
                        dataGridView_ref.Rows[idx1].Cells[0].Style.BackColor = Color.FromArgb(255, 219, 88);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Unselect_cell_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_ref.RowCount > 0)
                {
                    int idx1 = dataGridView_ref.CurrentCell.RowIndex;
                    if (idx1 >= 0)
                    {
                        dataGridView_ref.Rows[idx1].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Regular);
                        dataGridView_ref.Rows[idx1].Cells[0].Style.ForeColor = Color.White;
                        dataGridView_ref.Rows[idx1].Cells[0].Style.BackColor = Color.FromArgb(37, 37, 38);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Unselect_all_cells_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (dataGridView_ref.RowCount > 0)
                {
                    for (int j = 0; j < dataGridView_ref.Rows.Count; j++)
                    {
                        dataGridView_ref.Rows[j].Cells[0].Style.Font = new System.Drawing.Font(dataGridView_ref.Font, FontStyle.Regular);
                        dataGridView_ref.Rows[j].Cells[0].Style.ForeColor = Color.White;
                        dataGridView_ref.Rows[j].Cells[0].Style.BackColor = Color.FromArgb(37, 37, 38);
                    }

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }


        private void button_generate_excel_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_ref.RowCount > 0)
                {

                    System.Data.DataTable dt1 = new System.Data.DataTable();
                    dt1.Columns.Add("REF", typeof(string));

                    for (int j = 0; j < dataGridView_ref.Rows.Count; j++)
                    {

                        if (dataGridView_ref.Rows[j].Cells[0].Style.ForeColor == Color.FromArgb(0, 0, 0))
                        {
                            if (dt1.Rows.Count == 0)
                            {
                                dt1.Rows.Add();
                            }

                            dt1.Rows[dt1.Rows.Count - 1][0] = dataGridView_ref.Rows[j].Cells[0].Value;

                            string descr1 = Convert.ToString(dataGridView_ref.Rows[j].Cells[1].Value);
                            char split1 = Convert.ToChar('|');
                            string[] descr2 = descr1.Split(split1);
                            for (int k = 0; k < descr2.Length; k++)
                            {
                                dt1.Rows.Add();
                                dt1.Rows[dt1.Rows.Count - 1][0] = descr2[k];
                                dt1.Rows.Add();
                            }
                        }

                    }

                    string nume1 = System.DateTime.Now.Hour + "-" + System.DateTime.Now.Minute;
                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1);

                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void button_dan_Click(object sender, EventArgs e)
        {




            if (Display_dataTable != null && Display_dataTable.Rows.Count > 0)
            {
                drawing_list = new List<string>();
                for (int i = 0; i < Display_dataTable.Rows.Count; ++i)
                {
                    drawing_list.Add(Convert.ToString(Display_dataTable.Rows[i][0]));
                }




                Display_dataTable = new System.Data.DataTable();
                Display_dataTable.Columns.Add("Drawing", typeof(string));
                Display_dataTable.Columns.Add("Block Deleted", typeof(string));
                Display_dataTable.Columns.Add("Blocks Replaced", typeof(string));
                Display_dataTable.Columns.Add("Insert New Block", typeof(string));





                try
                {
                    set_enable_false();
                    _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                    if (Ag != null)
                    {
                        try
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


                                if (drawing_list != null && drawing_list.Count > 0)
                                {
                                    for (int i = 0; i < drawing_list.Count; ++i)
                                    {



                                        string file1 = drawing_list[i];





                                        if (System.IO.File.Exists(file1) == true)
                                        {
                                            Display_dataTable.Rows.Add();
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][0] = file1;
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][1] = "NO";
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][2] = "NO";
                                            Display_dataTable.Rows[Display_dataTable.Rows.Count - 1][3] = "NO";

                                            bool is_opened = false;


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
                                                        BlockTable BlockTable1 = Database2.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                                                        DBDictionary Layoutdict = (DBDictionary)Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead);

                                                        LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;


                                                        foreach (DBDictionaryEntry entry in Layoutdict)
                                                        {
                                                            Layout Layout0 = (Layout)Trans2.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForRead);
                                                            if (Layout0.TabOrder > 0)
                                                            {
                                                                BlockTableRecord BTrecord_PS = Trans2.GetObject(Layout0.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                                foreach (ObjectId id1 in BTrecord_PS)
                                                                {
                                                                    BlockReference bl1 = Trans2.GetObject(id1, OpenMode.ForWrite) as BlockReference;
                                                                    if (bl1 != null)
                                                                    {
                                                                        string nume_block = Functions.get_block_name_another_database(bl1, Database2);
                                                                        if (bl1.Layer == "0")
                                                                        {

                                                                            if (nume_block == "TEST" || nume_block == "TBLK_ATTRIBUTES")
                                                                            {
                                                                                bl1.Erase();
                                                                            }
                                                                        }
                                                                    }
                                                                    //DBText txt1 = Trans2.GetObject(id1, OpenMode.ForWrite) as DBText;
                                                                    //if (txt1 != null)
                                                                    //{
                                                                    //    if (txt1.Layer == "BORDERTXT")
                                                                    //    {
                                                                    //        txt1.Erase();
                                                                    //    }
                                                                    //}
                                                                    //MText Mtxt1 = Trans2.GetObject(id1, OpenMode.ForWrite) as MText;
                                                                    //if (Mtxt1 != null)
                                                                    //{
                                                                    //    if (Mtxt1.Layer == "MatchlineText" || Mtxt1.Layer == "TEXT_TOWNSHIP")
                                                                    //    {
                                                                    //        Mtxt1.Erase();
                                                                    //    }
                                                                    //}
                                                                    //Polyline poly1 = Trans2.GetObject(id1, OpenMode.ForWrite) as Polyline;
                                                                    //if (poly1 != null)
                                                                    //{
                                                                    //    if (poly1.Layer == "MatchLine")
                                                                    //    {
                                                                    //        poly1.Erase();
                                                                    //    }
                                                                    //}

                                                                }

                                                            }

                                                        }



                                                        Trans2.Commit();
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
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);

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

        }
    }
}
