using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using AcRx = Autodesk.AutoCAD.Runtime;
using System.Collections;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Globalization;
using acad = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.EditorInput;
using System.Data;

namespace Alignment_mdi
{
    public partial class Layer_controller_form : Form
    {




        System.Data.DataTable dt_dwg_selected = null;
        System.Data.DataTable dt_dwg_xl = null;
        System.Data.DataTable dt_stnd = null;
        System.Data.DataTable dt_tabs = null;


        List<string> drawing_list = null;

        private ContextMenuStrip ContextMenuStrip_open_alignment;

        public string stnd_xl_filename = "";


        public Layer_controller_form()
        {
            InitializeComponent();
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
            toolTip1.SetToolTip(button_dwg_to_excel, "Transfer drawings to the standard excel file");

        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_excel_to_dwg);
            lista_butoane.Add(button_load_block_attributes_to_excel);
            lista_butoane.Add(button_open_excel_tblk_attrib);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_select_excel_file);
            lista_butoane.Add(button_export_layers_from_selection);
            lista_butoane.Add(button_dwg_to_excel);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_excel_to_dwg);
            lista_butoane.Add(button_load_block_attributes_to_excel);
            lista_butoane.Add(button_open_excel_tblk_attrib);
            lista_butoane.Add(button_select_drawings);
            lista_butoane.Add(button_select_excel_file);
            lista_butoane.Add(button_export_layers_from_selection);
            lista_butoane.Add(button_dwg_to_excel);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
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


                    label_excel_to_green(System.IO.Path.GetFileName(File1));
                    stnd_xl_filename = File1;
                    Load_tab_names_to_combobox(comboBox_config_tabs);
                }
                else
                {
                    label_excel_to_red();
                }
            }
        }

        public void label_excel_to_red()
        {
            label_excel_file.Text = "Standard excel file not specified";
            label_excel_file.ForeColor = Color.Red;
            stnd_xl_filename = "";
        }

        private void label_excel_to_green(string file)
        {
            label_excel_file.Text = "Standards excel file: " + file;
            label_excel_file.ForeColor = Color.LimeGreen;
        }

        private void Load_tab_names_to_combobox(ComboBox combo1)
        {
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
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
            bool close_wkbook = true;

            try
            {
                if (stnd_xl_filename != "")
                {
                    foreach (Microsoft.Office.Interop.Excel.Workbook wbk in Excel1.Workbooks)
                    {
                        if (wbk.Name == System.IO.Path.GetFileName(stnd_xl_filename))
                        {
                            close_wkbook = false;
                            Workbook1 = wbk;
                        }
                    }
                    if (close_wkbook == true)
                    {
                        Workbook1 = Excel1.Workbooks.Open(stnd_xl_filename);
                    }
                    combo1.Items.Clear();
                    dt_tabs = new System.Data.DataTable();
                    dt_tabs.Columns.Add("Stnd", typeof(string));
                    dt_tabs.Columns.Add("SourceDrawing", typeof(string));
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                        {
                            //if (W1.Name.ToUpper().Contains("STND") == true)
                            //{
                            combo1.Items.Add(W1.Name);
                            dt_tabs.Rows.Add();
                            dt_tabs.Rows[dt_tabs.Rows.Count - 1][0] = W1.Name;
                            dt_tabs.Rows[dt_tabs.Rows.Count - 1][1] = W1.Range["B1"].Value2;
                            //}
                        }
                        if (close_wkbook == true)
                        {
                            Workbook1.Close();
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
                    finally
                    {
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }
                    combo1.Items.Insert(0, "New STND Tab");
                    combo1.SelectedIndex = 0;
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }


        private void Create_layer_controller_spreadsheet_Click(object sender, EventArgs e)
        {
           

            string file1 = System.IO.Path.GetFileName(stnd_xl_filename);

            if (Functions.Get_if_workbook_is_open_in_Excel(file1) == true)
            {
                MessageBox.Show("Please close the layer controller file");
                return;
            }

            string dwg1 = "";

            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Autocad files (*.dwg)|*.dwg";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    dwg1 = fbd.FileName;
                }
                else
                {
                    return;
                }
            }

            set_enable_false();

            try
            {
                List<string> lista1 = new List<string>();
                lista1.Add(dwg1);
                System.Data.DataTable dt1 = load_layers_from_dwg(lista1, false);
                string tab_name = comboBox_config_tabs.Text;
                if (comboBox_config_tabs.SelectedIndex == 0) tab_name = "";
                Transfer_datatable_to_existing_excel_spreadsheet_layer_controller_config(dt1, stnd_xl_filename, dwg1, tab_name);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            set_enable_true();
        }

        private System.Data.DataTable load_layers_from_dwg(List<string> lista1, bool extra_columns)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();


            System.Data.DataTable dt3 = new System.Data.DataTable();
            dt3.Columns.Add("Layer Name", typeof(string));
            dt3.Columns.Add("ON/OFF", typeof(string));
            dt3.Columns.Add("THAW/FREEZE", typeof(string));
            dt3.Columns.Add("Color", typeof(string));
            dt3.Columns.Add("Linetype", typeof(string));
            dt3.Columns.Add("Lineweight", typeof(string));
            dt3.Columns.Add("Transparency", typeof(string));
            dt3.Columns.Add("Plot", typeof(string));

            if (extra_columns == true)
            {
                dt3.Columns.Add("DWG", typeof(string));
            }

            if (lista1 == null || lista1.Count == 0)
            {
                return dt3;
            }

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    for (int i = 0; i < lista1.Count; ++i)
                    {
                        string dwg1 = lista1[i];
                        if (System.IO.File.Exists(dwg1) == true)
                        {
                            using (Database Database2 = new Database(false, true))
                            {
                                Database2.ReadDwgFile(dwg1, FileOpenMode.OpenForReadAndAllShare, true, "");
                                //System.IO.FileShare.ReadWrite, false, null);
                                Database2.CloseInput(true);
                                Database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);
                                HostApplicationServices.WorkingDatabase = Database2;
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                {
                                    LayerTable LayerTable2 = Trans2.GetObject(Database2.LayerTableId, OpenMode.ForRead) as LayerTable;
                                    foreach (ObjectId id1 in LayerTable2)
                                    {
                                        LayerTableRecord ltr = Trans2.GetObject(id1, OpenMode.ForRead) as LayerTableRecord;
                                        if (ltr != null)
                                        {
                                            dt3.Rows.Add();
                                            dt3.Rows[dt3.Rows.Count - 1][0] = ltr.Name;
                                            if (ltr.IsOff == false)
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][1] = "ON";
                                            }
                                            else
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][1] = "OFF";
                                            }

                                            if (ltr.IsFrozen == false)
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][2] = "THAWED";
                                            }
                                            else
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][2] = "FROZEN";
                                            }

                                            Autodesk.AutoCAD.Colors.Color color1 = ltr.Color;
                                            if (color1.IsByAci == true)
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][3] = color1.ColorIndex;
                                            }
                                            else
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][3] = color1.Red + "," + color1.Green + "," + color1.Blue;
                                            }





                                            LinetypeTableRecord linetype1 = Trans2.GetObject(ltr.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                            if (linetype1 != null)
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1][4] = linetype1.Name;

                                                //if(dt1.Columns.Contains("asciidescription")==false)
                                                //{
                                                //    dt1.Columns.Add("asciidescription", typeof(string));
                                                //}
                                                //if (dt1.Columns.Contains("patternlength") == false)
                                                //{
                                                //    dt1.Columns.Add("patternlength", typeof(double));
                                                //}
                                                //if (dt1.Columns.Contains("numdashes") == false)
                                                //{
                                                //    dt1.Columns.Add("numdashes", typeof(int));
                                                //}


                                                //dt1.Rows.Add();
                                                //dt1.Rows[dt1.Rows.Count - 1]["asciidescription"] = linetype1.AsciiDescription;
                                                //dt1.Rows[dt1.Rows.Count - 1]["patternlength"] = linetype1.PatternLength;
                                                //dt1.Rows[dt1.Rows.Count - 1]["numdashes"] = linetype1.NumDashes;

                                                //for (int k=0; k<linetype1.NumDashes;++k)
                                                //{
                                                //    if (dt1.Columns.Contains("dasheslengthat"+ Convert.ToString(k)) == false)
                                                //    {
                                                //        dt1.Columns.Add("dasheslengthat" + Convert.ToString(k ), typeof(double));
                                                //    }

                                                //    dt1.Rows[dt1.Rows.Count - 1]["dasheslengthat" + Convert.ToString(k )] = linetype1.DashLengthAt(k);

                                                //    if (dt1.Columns.Contains("textat" + Convert.ToString(k)) == false)
                                                //    {
                                                //        dt1.Columns.Add("textat" + Convert.ToString(k), typeof(string));
                                                //    }


                                                //    if(linetype1.TextAt(k)!= "")
                                                //    {
                                                //        dt1.Rows[dt1.Rows.Count - 1]["textat" + Convert.ToString(k)] = linetype1.TextAt(k);

                                                //    }

                                                //}



                                                //ltr.Name = "COLD_WATER_SUPPLY";
                                                //ltr.AsciiDescription =
                                                //    "Cold water supply ---- CW ---- CW ---- CW ----";
                                                //ltr.PatternLength = 0.9;
                                                //ltr.NumDashes = 3;
                                                //// Dash #1
                                                //ltr.SetDashLengthAt(0, 0.5);
                                                //// Dash #2
                                                //ltr.SetDashLengthAt(1, -0.2);
                                                //ltr.SetShapeStyleAt(1, tt["Standard"]);
                                                //ltr.SetShapeNumberAt(1, 0);
                                                //ltr.SetShapeScaleAt(1, 0.1);
                                                //ltr.SetTextAt(1, "CW");
                                                //ltr.SetShapeRotationAt(1, 0);
                                                //ltr.SetShapeOffsetAt(1, new Vector2d(0, -0.05));
                                                //// Dash #3
                                                //ltr.SetDashLengthAt(2, -0.2);
                                            }

                                            dt3.Rows[dt3.Rows.Count - 1][5] = ltr.LineWeight;
                                            dt3.Rows[dt3.Rows.Count - 1][6] = Convert.ToString(ltr.Transparency).Replace("(", "").Replace(")", "");

                                            string yesno = "YES";
                                            if (ltr.IsPlottable == false) yesno = "NO";

                                            dt3.Rows[dt3.Rows.Count - 1][7] = yesno;

                                            if (extra_columns == true)
                                            {
                                                dt3.Rows[dt3.Rows.Count - 1]["DWG"] = dwg1;
                                            }
                                        }
                                    }
                                }
                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                            }
                        }
                    }
                }
            }
            return dt3;
        }

        void Transfer_datatable_to_existing_excel_spreadsheet_layer_controller_config(System.Data.DataTable dt1, string excel_filename, string dwg_filename, string sheetname)
        {
            if (dt1 != null && dt1.Rows.Count > 0)
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

                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(excel_filename);

                    List<string> lista_pages = new List<string>();

                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                    {
                        if (Wx.Name.ToUpper() == sheetname.ToUpper())
                        {
                            W1 = Wx;
                        }
                        lista_pages.Add(Wx.Name);
                    }

                    if (sheetname == "")
                    {
                        W1 = Workbook1.Worksheets.Add(Before: Workbook1.Worksheets[1]);
                        int conf_index = 1;
                        bool named = false;
                        do
                        {
                            if (lista_pages.Contains("STND_" + Convert.ToString(conf_index)) == false)
                            {
                                W1.Name = "STND_" + Convert.ToString(conf_index);
                                sheetname = W1.Name;
                                named = true;
                            }
                            else
                            {
                                ++conf_index;
                            }

                        } while (named == false);
                    }


                    if (W1 != null)
                    {
                        int maxRows = dt1.Rows.Count;
                        int maxCols = dt1.Columns.Count;
                        char col1 = (char)(64 + maxCols);
                        W1.Range["D:D"].NumberFormat = "@";

                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A4:" + col1 + Convert.ToString(4 + maxRows - 1)];
                        range1.ClearContents();
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
                            W1.Cells[3, i + 1].value2 = dt1.Columns[i].ColumnName.ToUpper();
                        }
                        range1.Value2 = values1;
                        W1.Columns["A"].ColumnWidth = 50;
                        W1.Columns["B:H"].ColumnWidth = 22;
                        W1.Range["A2:H2"].Merge();
                        W1.Range["A2"].Value2 = "OVERALL LAYER SETTINGS";
                        W1.Range["A1"].Value2 = "SOURCE DRAWING";
                        W1.Range["B1"].Value2 = dwg_filename;
                        Workbook1.Save();
                    }

                    dt_tabs.Rows.Clear();
                    comboBox_config_tabs.Items.Clear();
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                    {
                        if (W2.Name.ToUpper().Contains("STND") == true)
                        {
                            comboBox_config_tabs.Items.Add(W2.Name);
                            dt_tabs.Rows.Add();
                            dt_tabs.Rows[dt_tabs.Rows.Count - 1][0] = W2.Name;
                            dt_tabs.Rows[dt_tabs.Rows.Count - 1][1] = W2.Range["B1"].Value2;
                        }
                    }

                    comboBox_config_tabs.SelectedIndex = comboBox_config_tabs.Items.IndexOf(sheetname);
                    label_dwg.Text = Workbook1.Worksheets[sheetname].Range["B1"].Value2;
                    try
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
                    finally
                    {
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }

        private System.Drawing.Color AcadColorAciToDrawingColor(Autodesk.AutoCAD.Colors.Color color)
        {
            byte aci = Convert.ToByte(color.ColorIndex);
            int aRGB = Autodesk.AutoCAD.Colors.EntityColor.LookUpRgb(aci);
            byte[] ch = BitConverter.GetBytes(aRGB);
            if (!BitConverter.IsLittleEndian)
            {
                Array.Reverse(ch);
            }
            int r = ch[2];
            int g = ch[1];
            int b = ch[0];

            return System.Drawing.Color.FromArgb(r, g, b);
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
            dt_dwg_selected = null;

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
                    if (dt_dwg_selected != null)
                    {
                        if (dt_dwg_selected.Rows.Count - 1 >= Index1)
                        {
                            string fisier_generat = dt_dwg_selected.Rows[Index1][0].ToString();

                            string Output_folder = "";
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







        private void button_open_excel_tblk_attributes_Click(object sender, EventArgs e)
        {

            {

                try
                {




                    if (System.IO.File.Exists(stnd_xl_filename) == false)
                    {

                        MessageBox.Show("the STANDARDS EXCEL FILE data file does not exist");
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
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(stnd_xl_filename);




                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);

                }


            }
        }

        private void button_select_drawings_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "Autocad files (*.dwg)|*.dwg";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (dt_dwg_selected == null)
                    {
                        dt_dwg_selected = new System.Data.DataTable();
                        dt_dwg_selected.Columns.Add("Drawing", typeof(string));
                    }

                    if (drawing_list == null)
                    {
                        drawing_list = new List<string>();
                    }


                    //Layer_controller_form.tpage_processing.Show();

                    foreach (string file1 in fbd.FileNames)
                    {
                        if (drawing_list.Contains(file1) == false)
                        {
                            dt_dwg_selected.Rows.Add();
                            dt_dwg_selected.Rows[dt_dwg_selected.Rows.Count - 1][0] = file1;
                            drawing_list.Add(file1);
                        }
                    }

                    dt_dwg_selected = Functions.Sort_data_table(dt_dwg_selected, "Drawing");

                    dataGridView_drawings.DataSource = dt_dwg_selected;
                    dataGridView_drawings.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_drawings.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    dataGridView_drawings.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                    dataGridView_drawings.DefaultCellStyle.ForeColor = Color.White;
                    dataGridView_drawings.EnableHeadersVisualStyles = false;

                    // Layer_controller_form.tpage_processing.Hide();
                    this.MdiParent.WindowState = FormWindowState.Normal;
                }
            }
        }





        private void button_compare_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Layer Name", typeof(string));
            dt1.Columns.Add("ON/OFF", typeof(string));
            dt1.Columns.Add("THAW/FREEZE", typeof(string));
            dt1.Columns.Add("Color", typeof(string));
            dt1.Columns.Add("Linetype", typeof(string));
            dt1.Columns.Add("Lineweight", typeof(string));
            dt1.Columns.Add("Transparency", typeof(string));
            dt1.Columns.Add("Plot", typeof(string));
            dt1.Columns.Add("Stnd", typeof(string));
            dt1.Columns.Add("Config_type", typeof(string));

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Columns.Add("Drawing", typeof(string));
            dt2.Columns.Add("Stnd", typeof(string));

            System.Data.DataTable dt3 = null;

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
                if (stnd_xl_filename != "")
                {
                    bool close_wkbook = true;
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                    foreach (Microsoft.Office.Interop.Excel.Workbook wbk in Excel1.Workbooks)
                    {
                        if (wbk.Name == System.IO.Path.GetFileName(stnd_xl_filename))
                        {
                            close_wkbook = false;
                            Workbook1 = wbk;
                        }
                    }

                    if (close_wkbook == true)
                    {
                        Workbook1 = Excel1.Workbooks.Open(stnd_xl_filename);
                    }


                    bool exista_tab_dwgs = false;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet w1 in Workbook1.Worksheets)
                    {
                        if (w1.Name.ToUpper() == "DWGS")
                        {
                            exista_tab_dwgs = true;
                        }
                    }
                    dt2 = load_from_excel_list_of_drawings(Workbook1, dt2);
                    if (exista_tab_dwgs == true && dt2.Rows.Count > 0)
                    {
                        dt1 = load_from_Excel_layer_data(Workbook1, dt1);


                        List<string> lista2 = new List<string>();
                        for (int j = 0; j < dt2.Rows.Count; ++j)
                        {
                            if (dt2.Rows[j][0] != DBNull.Value)
                            {
                                lista2.Add(Convert.ToString(dt2.Rows[j][0]));
                            }
                        }
                        dt3 = load_layers_from_dwg(lista2, true);
                    }

                    try
                    {
                        if (close_wkbook == true)
                        {
                            Workbook1.Close();

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
                    finally
                    {
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }


                    if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0 && dt3.Rows.Count > 0)
                    {

                        System.Data.DataTable dt4 = new System.Data.DataTable();
                        dt4.Columns.Add("Layer Name", typeof(string));
                        dt4.Columns.Add("ON/OFF", typeof(string));
                        dt4.Columns.Add("THAW/FREEZE", typeof(string));
                        dt4.Columns.Add("Color", typeof(string));
                        dt4.Columns.Add("Linetype", typeof(string));
                        dt4.Columns.Add("Lineweight", typeof(string));
                        dt4.Columns.Add("Transparency", typeof(string));
                        dt4.Columns.Add("Plot", typeof(string));
                        dt4.Columns.Add("Error Description", typeof(string));
                        dt4.Columns.Add("DWG", typeof(string));
                        dt4.Columns.Add("STND Tab", typeof(string));

                        dt4.Columns.Add("Value listed in STND Tab", typeof(string));

                        int row_start = 0;


                        for (int j = 0; j < dt2.Rows.Count; ++j)
                        {
                            bool is_error = false;

                            string dwg2 = "";
                            string tab2 = "";

                            if (dt2.Rows[j][0] != DBNull.Value)
                            {
                                dwg2 = Convert.ToString(dt2.Rows[j][0]);
                            }

                            if (dt2.Rows[j][1] != DBNull.Value)
                            {
                                tab2 = Convert.ToString(dt2.Rows[j][1]);
                            }

                            string dwg22 = System.IO.Path.GetFileNameWithoutExtension(dwg2);

                            System.Data.DataTable dt33 = new System.Data.DataTable();
                            dt33 = dt3.Clone();

                            for (int k = 0; k < dt3.Rows.Count; ++k)
                            {
                                string dwg3 = "";
                                if (dt3.Rows[k]["DWG"] != DBNull.Value)
                                {
                                    dwg3 = Convert.ToString(dt3.Rows[k]["DWG"]);
                                    if (dwg3.ToLower() == dwg2.ToLower())
                                    {
                                        dt33.ImportRow(dt3.Rows[k]);
                                    }
                                }
                            }



                            System.Data.DataTable dt11 = new System.Data.DataTable();
                            dt11 = dt1.Clone();

                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                string tab1 = "";
                                if (dt1.Rows[i]["Stnd"] != DBNull.Value)
                                {
                                    tab1 = Convert.ToString(dt1.Rows[i]["Stnd"]);
                                    if (tab1.ToLower() == tab2.ToLower())
                                    {
                                        dt11.ImportRow(dt1.Rows[i]);
                                    }
                                }
                            }

                            dt11.Columns.RemoveAt(dt11.Columns.Count - 1);
                            dt11.Columns.RemoveAt(dt11.Columns.Count - 1);

                            dt11.TableName = "11";
                            dt33.TableName = "33";
                            DataSet dataset1 = new DataSet();
                            dataset1.Tables.Add(dt11);
                            dataset1.Tables.Add(dt33);

                            DataRelation relation0 = new DataRelation("xxx", dt11.Columns[0], dt33.Columns[0], false);
                            dataset1.Relations.Add(relation0);
                            DataRelation relation1 = new DataRelation("xyz", dt33.Columns[0], dt11.Columns[0], false);
                            dataset1.Relations.Add(relation1);



                            for (int n = 0; n < dt11.Rows.Count; ++n)
                            {
                                if (dt11.Rows[n].GetChildRows(relation0).Length == 0)
                                {
                                    dt4.Rows.Add();
                                    dt4.Rows[dt4.Rows.Count - 1][0] = dt11.Rows[n][0];
                                    dt4.Rows[dt4.Rows.Count - 1][1] = dt11.Rows[n][1];
                                    dt4.Rows[dt4.Rows.Count - 1][2] = dt11.Rows[n][2];
                                    dt4.Rows[dt4.Rows.Count - 1][3] = dt11.Rows[n][3];
                                    dt4.Rows[dt4.Rows.Count - 1][4] = dt11.Rows[n][4];
                                    dt4.Rows[dt4.Rows.Count - 1][5] = dt11.Rows[n][5];
                                    dt4.Rows[dt4.Rows.Count - 1][6] = dt11.Rows[n][6];
                                    dt4.Rows[dt4.Rows.Count - 1][7] = dt11.Rows[n][7];
                                    dt4.Rows[dt4.Rows.Count - 1][8] = "Layer not found in DWG, listed in STDN tab";
                                    dt4.Rows[dt4.Rows.Count - 1][9] = dwg22;
                                    dt4.Rows[dt4.Rows.Count - 1][10] = tab2;
                                    is_error = true;
                                }
                                else
                                {
                                    for (int k = 1; k < 8; ++k)
                                    {
                                        string dwg_layer_val = Convert.ToString(dt11.Rows[n].GetChildRows(relation0)[0][k]);
                                        string excel_layer_val = Convert.ToString(dt11.Rows[n][k]);

                                        if (dwg_layer_val.ToLower() != excel_layer_val.ToLower())
                                        {
                                            dt4.Rows.Add();
                                            dt4.Rows[dt4.Rows.Count - 1][0] = dt11.Rows[n].GetChildRows(relation0)[0][0];
                                            dt4.Rows[dt4.Rows.Count - 1][1] = dt11.Rows[n].GetChildRows(relation0)[0][1];
                                            dt4.Rows[dt4.Rows.Count - 1][2] = dt11.Rows[n].GetChildRows(relation0)[0][2];
                                            dt4.Rows[dt4.Rows.Count - 1][3] = dt11.Rows[n].GetChildRows(relation0)[0][3];
                                            dt4.Rows[dt4.Rows.Count - 1][4] = dt11.Rows[n].GetChildRows(relation0)[0][4];
                                            dt4.Rows[dt4.Rows.Count - 1][5] = dt11.Rows[n].GetChildRows(relation0)[0][5];
                                            dt4.Rows[dt4.Rows.Count - 1][6] = dt11.Rows[n].GetChildRows(relation0)[0][6];
                                            dt4.Rows[dt4.Rows.Count - 1][7] = dt11.Rows[n].GetChildRows(relation0)[0][7];

                                            if (k == 1)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer on/off missmatch";
                                            }
                                            if (k == 2)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer thaw/freeze missmatch";
                                            }
                                            if (k == 3)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer color missmatch";
                                            }
                                            if (k == 4)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer linetype missmatch";
                                            }
                                            if (k == 5)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer lineweight missmatch";
                                            }
                                            if (k == 6)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer transparency missmatch";
                                            }

                                            if (k == 7)
                                            {
                                                dt4.Rows[dt4.Rows.Count - 1][8] = "Layer plot yes/no missmatch";
                                            }

                                            dt4.Rows[dt4.Rows.Count - 1][9] = dwg22;
                                            dt4.Rows[dt4.Rows.Count - 1][10] = tab2;
                                            dt4.Rows[dt4.Rows.Count - 1][11] = dt11.Rows[n][k];
                                            is_error = true;

                                        }
                                    }
                                }
                            }

                            for (int l = 0; l < dt33.Rows.Count; ++l)
                            {
                                if (dt33.Rows[l].GetChildRows(relation1).Length == 0)
                                {
                                    dt4.Rows.Add();
                                    dt4.Rows[dt4.Rows.Count - 1][0] = dt33.Rows[l][0];
                                    dt4.Rows[dt4.Rows.Count - 1][1] = dt33.Rows[l][1];
                                    dt4.Rows[dt4.Rows.Count - 1][2] = dt33.Rows[l][2];
                                    dt4.Rows[dt4.Rows.Count - 1][3] = dt33.Rows[l][3];
                                    dt4.Rows[dt4.Rows.Count - 1][4] = dt33.Rows[l][4];
                                    dt4.Rows[dt4.Rows.Count - 1][5] = dt33.Rows[l][5];
                                    dt4.Rows[dt4.Rows.Count - 1][6] = dt33.Rows[l][6];
                                    dt4.Rows[dt4.Rows.Count - 1][7] = dt33.Rows[l][7];
                                    dt4.Rows[dt4.Rows.Count - 1][8] = "Extra Layer found in DWG,  NOT listed in STDN tab";
                                    dt4.Rows[dt4.Rows.Count - 1][9] = dwg22;
                                    dt4.Rows[dt4.Rows.Count - 1][10] = tab2;
                                    is_error = true;

                                }
                            }

                            System.Data.DataRow row0 = dt4.NewRow();
                            System.Data.DataRow row00 = dt4.NewRow();
                            System.Data.DataRow row1 = dt4.NewRow();
                            row1[0] = dwg2 + " - " + tab2;
                            System.Data.DataRow row2 = dt4.NewRow();
                            row2[0] = "OVERALL LAYER SETTINGS";

                            System.Data.DataRow row3 = dt4.NewRow();
                            for (int m = 0; m < dt4.Columns.Count; ++m)
                            {
                                row3[m] = dt4.Columns[m].ColumnName;
                            }

                            if (is_error == true)
                            {

                                dt4.Rows.InsertAt(row3, row_start);
                                dt4.Rows.InsertAt(row2, row_start);
                                dt4.Rows.InsertAt(row1, row_start);
                                if (j > 0)
                                {
                                    dt4.Rows.InsertAt(row0, row_start);
                                    dt4.Rows.InsertAt(row00, row_start);
                                }
                            }

                            is_error = false;
                            row_start = dt4.Rows.Count;

                            dataset1.Relations.Remove(relation0);
                            dataset1.Relations.Remove(relation1);


                            dataset1.Tables.Remove(dt11);
                            dataset1.Tables.Remove(dt33);
                            dataset1.Dispose();


                        }
                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt4);
                    }


                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable load_from_Excel_layer_data(Microsoft.Office.Interop.Excel.Workbook Workbook1, System.Data.DataTable dt1)
        {

            foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
            {

                if (W1.Range["A2"].Value2 == "OVERALL LAYER SETTINGS")
                {
                    object[,] values1 = new object[50001, 1];
                    values1 = W1.Range["A4" + ":A50004"].Value2;
                    int EndRow = 0;
                    for (int i = 1; i <= values1.Length; ++i)
                    {
                        object Val1 = values1[i, 1];
                        if (Val1 != null)
                        {
                            if (Val1.ToString().Contains("CONTROLLER VIEWPORT") == true)
                            {
                                EndRow = i + 4 - 1 - 1;
                                i = values1.Length + 1;
                            }
                        }
                        else
                        {
                            EndRow = i + 4 - 1 - 1;
                            i = values1.Length + 1;
                        }
                    }
                    object[,] values_AH = new object[EndRow - 4 + 1, 11];
                    values_AH = W1.Range["A4:K" + EndRow.ToString()].Value2;
                    for (int i = 1; i <= EndRow - 4 + 1; ++i)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1]["Stnd"] = W1.Name;
                        dt1.Rows[dt1.Rows.Count - 1]["Config_type"] = "OVERALL";

                        for (int j = 0; j < 11; ++j)
                        {
                            object Valoare = values_AH[i, j + 1];
                            if (Valoare == null) Valoare = DBNull.Value;
                            dt1.Rows[dt1.Rows.Count - 1][j] = Valoare;
                        }
                    }
                }

            }
            return dt1;
        }

        private System.Data.DataTable load_from_excel_list_of_drawings(Microsoft.Office.Interop.Excel.Workbook Workbook1, System.Data.DataTable dt2)
        {

            foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
            {
                if (W1.Name.ToUpper() == "DWGS")
                {
                    object[,] values1 = new object[50001, 1];
                    values1 = W1.Range["A2:A50002"].Value2;
                    int EndRow = 0;
                    for (int i = 1; i <= values1.Length; ++i)
                    {
                        object Val1 = values1[i, 1];
                        if (Val1 != null)
                        {
                            if (Val1.ToString().Replace(" ", "") == "")
                            {
                                EndRow = i + 2 - 1 - 1;
                                i = values1.Length + 1;
                            }
                        }
                        else
                        {
                            EndRow = i + 2 - 1 - 1;
                            i = values1.Length + 1;
                        }
                    }
                    object[,] values_AB = new object[EndRow - 2 + 1, 2];
                    values_AB = W1.Range["A2:B" + EndRow.ToString()].Value2;
                    for (int i = 1; i <= EndRow - 2 + 1; ++i)
                    {
                        dt2.Rows.Add();
                        for (int j = 0; j < 2; ++j)
                        {
                            object Valoare = values_AB[i, j + 1];
                            if (Valoare == null) Valoare = DBNull.Value;
                            dt2.Rows[dt2.Rows.Count - 1][j] = Valoare;
                        }
                    }
                }
            }
            return dt2;
        }

        private void ComboBox_config_tabs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dt_tabs != null && dt_tabs.Rows.Count > 0)
            {
                if (comboBox_config_tabs.SelectedIndex == 0)
                {
                    label_dwg.Text = "DWG:";
                }
                else
                {
                    string txt1 = comboBox_config_tabs.Text;
                    for (int i = 0; i < dt_tabs.Rows.Count; ++i)
                    {
                        if (dt_tabs.Rows[i][0] != DBNull.Value && dt_tabs.Rows[i][1] != DBNull.Value)
                        {
                            string txt2 = Convert.ToString(dt_tabs.Rows[i][0]);
                            string dwg2 = Convert.ToString(dt_tabs.Rows[i][1]);
                            if (txt1 == txt2)
                            {
                                if (dwg2.Replace(" ", "") != "")
                                {
                                    label_dwg.Text = "DWG: " + dwg2;
                                }
                                else
                                {
                                    label_dwg.Text = "DWG:";
                                }
                            }
                        }
                    }

                }
            }


        }

        private void Button_dwg_to_excel_Click(object sender, EventArgs e)
        {
            Transfer_drawing_list_to_layer_controller_config(drawing_list, stnd_xl_filename);

        }

        private void Transfer_drawing_list_to_layer_controller_config(List<string> lista1, string excel_filename)
        {
            if (lista1 != null && lista1.Count > 0)
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

                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(excel_filename);
                    try
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook1.Worksheets)
                        {
                            if (Wx.Name.ToUpper() == "DWGS")
                            {
                                W1 = Wx;
                            }

                        }

                        if (W1 == null)
                        {
                            W1 = Workbook1.Worksheets.Add(Before: Workbook1.Worksheets[1]);
                            W1.Name = "DWGS";
                        }


                        if (W1 != null)
                        {
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:C" + Convert.ToString(drawing_list.Count + 1)];
                            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["A:C"];
                            range2.ClearContents();
                            object[,] values1 = new object[drawing_list.Count + 1, 3];

                            values1[0, 0] = "DRAWING NAME";
                            values1[0, 1] = "STANDARD";
                            values1[0, 2] = "NOTE";
                            for (int i = 0; i < drawing_list.Count; ++i)
                            {
                                values1[i + 1, 0] = lista1[i];
                            }
                            range1.Value2 = values1;
                            W1.Columns["A"].ColumnWidth = 63;
                            W1.Columns["B"].ColumnWidth = 15;
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
        }



        private void button_excel_to_dwg_Click(object sender, EventArgs e)
        {

            if (dt_dwg_selected != null && dt_dwg_selected.Rows.Count > 0)
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
                            load_from_excel_dwg_and_standards_data();

                            if (dt_stnd == null || dt_stnd.Rows.Count == 0)
                            {
                                MessageBox.Show("no stnd found");
                                set_enable_true();
                            }

                            if (dt_dwg_xl != null && dt_dwg_xl.Rows.Count > 0 && dt_stnd != null && dt_stnd.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt_dwg_selected.Rows.Count; ++i)
                                {
                                    if (dt_dwg_selected.Rows[i][0] != DBNull.Value)
                                    {
                                        string file1 = Convert.ToString(dt_dwg_selected.Rows[i][0]);
                                        if (System.IO.File.Exists(file1) == true)
                                        {
                                            for (int k = 0; k < dt_dwg_xl.Rows.Count; ++k)
                                            {
                                                if (dt_dwg_xl.Rows[k][0] != DBNull.Value)
                                                {
                                                    string file2 = Convert.ToString(dt_dwg_xl.Rows[k][0]);
                                                    string compare2 = file2;
                                                    if (file2.Contains("\\") == true)
                                                    {
                                                        compare2 = System.IO.Path.GetFileNameWithoutExtension(file2);
                                                    }
                                                    compare2 = compare2.Replace(".dwg", "");
                                                    if (System.IO.Path.GetFileNameWithoutExtension(file1) == compare2)
                                                    {
                                                        if (dt_dwg_xl.Rows[k]["Stnd"] != DBNull.Value)
                                                        {
                                                            string stnd1 = Convert.ToString(dt_dwg_xl.Rows[k]["Stnd"]);
                                                            System.Data.DataTable dt1 = new System.Data.DataTable();
                                                            dt1 = dt_stnd.Clone();
                                                            for (int j = 0; j < dt_stnd.Rows.Count; ++j)
                                                            {
                                                                if (dt_stnd.Rows[j]["Stnd"] != DBNull.Value)
                                                                {
                                                                    string stnd2 = Convert.ToString(dt_stnd.Rows[j]["Stnd"]);
                                                                    if (stnd1.ToLower() == stnd2.ToLower())
                                                                    {
                                                                        dt1.ImportRow(dt_stnd.Rows[j]);
                                                                    }
                                                                }
                                                            }
                                                            if (dt1.Rows.Count > 0)
                                                            {

                                                                write_stnd_to_dwg(file1, dt1);
                                                            }
                                                        }
                                                        k = dt_dwg_xl.Rows.Count;
                                                    }
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




        }

        private void write_stnd_to_dwg(string path1, System.Data.DataTable dt1)
        {
            List<string> list_lw = new List<string>();
            list_lw.Add("ByLineWeightDefault");
            list_lw.Add("LineWeight000");
            list_lw.Add("LineWeight005");
            list_lw.Add("LineWeight009");
            list_lw.Add("LineWeight013");
            list_lw.Add("LineWeight015");
            list_lw.Add("LineWeight018");
            list_lw.Add("LineWeight020");
            list_lw.Add("LineWeight025");
            list_lw.Add("LineWeight030");
            list_lw.Add("LineWeight035");
            list_lw.Add("LineWeight040");
            list_lw.Add("LineWeight050");
            list_lw.Add("LineWeight060");
            list_lw.Add("LineWeight070");
            list_lw.Add("LineWeight080");
            list_lw.Add("LineWeight090");
            list_lw.Add("LineWeight100");
            list_lw.Add("LineWeight106");
            list_lw.Add("LineWeight120");
            list_lw.Add("LineWeight140");
            list_lw.Add("LineWeight158");
            list_lw.Add("LineWeight200");
            list_lw.Add("LineWeight211");



            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    if (System.IO.File.Exists(path1) == true)
                    {
                        try
                        {
                            using (Database Database2 = new Database(false, true))
                            {
                                Database2.ReadDwgFile(path1, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                //System.IO.FileShare.ReadWrite, false, null);
                                Database2.CloseInput(true);
                                Database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);

                                HostApplicationServices.WorkingDatabase = Database2;

                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                {

                                    LayerTable LayerTable2 = Trans2.GetObject(Database2.LayerTableId, OpenMode.ForRead) as LayerTable;

                                    if (comboBox_visretain.Text == "0")
                                    {
                                        Database2.Visretain = false;
                                    }

                                    if (comboBox_visretain.Text == "1")
                                    {
                                        Database2.Visretain = true;
                                    }

                                    ObjectIdCollection viewportCollection = new ObjectIdCollection();
                                    DBDictionary dBLayoutDictionary = Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                                    foreach (DBDictionaryEntry entry in dBLayoutDictionary)
                                    {
                                        Layout layout = Trans2.GetObject(entry.Value, OpenMode.ForRead) as Layout;
                                        if (entry.Key != "Model")
                                        {
                                            viewportCollection = layout.GetViewports();
                                            viewportCollection.RemoveAt((0));
                                        }
                                    }


                                    foreach (ObjectId id1 in LayerTable2)
                                    {
                                        LayerTableRecord ltr = Trans2.GetObject(id1, OpenMode.ForWrite) as LayerTableRecord;
                                        if (ltr != null)
                                        {
                                            string nume1 = ltr.Name;

                                            bool proceseaza = true;
                                            if (nume1.Contains("|") == true && ltr.IsResolved == false)
                                            {
                                                proceseaza = false;
                                            }

                                            if (proceseaza == true)
                                            {
                                                for (int j = 0; j < dt1.Rows.Count; ++j)
                                                {
                                                    if (dt1.Rows[j][0] != DBNull.Value)
                                                    {
                                                        string nume2 = Convert.ToString(dt1.Rows[j][0]);
                                                        if (nume1.ToLower() == nume2.ToLower())
                                                        {
                                                            bool is_off = false;
                                                            if (dt1.Rows[j][1] != DBNull.Value)
                                                            {
                                                                string val2 = Convert.ToString(dt1.Rows[j][1]);
                                                                if (val2.ToLower().Contains("off") == true)
                                                                {
                                                                    is_off = true;
                                                                }
                                                            }
                                                            ltr.IsOff = is_off;

                                                            if (Database2.Clayer != ltr.ObjectId)
                                                            {
                                                                bool is_frozen = false;
                                                                if (dt1.Rows[j][2] != DBNull.Value)
                                                                {
                                                                    string val2 = Convert.ToString(dt1.Rows[j][2]);
                                                                    if (val2.ToLower().Contains("frozen") == true)
                                                                    {
                                                                        is_frozen = true;
                                                                    }
                                                                }
                                                                ltr.IsFrozen = is_frozen;
                                                            }


                                                            bool is_plottable = true;
                                                            if (dt1.Rows[j][7] != DBNull.Value)
                                                            {
                                                                string val2 = Convert.ToString(dt1.Rows[j][7]);
                                                                if (val2.ToLower().Contains("no") == true || val2.ToLower().Contains("false") == true)
                                                                {
                                                                    is_plottable = false;
                                                                }
                                                            }
                                                            ltr.IsPlottable = is_plottable;


                                                            if (dt1.Rows[j][3] != DBNull.Value)
                                                            {
                                                                string color1_string = Convert.ToString(dt1.Rows[j][3]);
                                                                if (color1_string.ToLower().Contains(",") == false && Functions.IsNumeric(color1_string) == true)
                                                                {
                                                                    Autodesk.AutoCAD.Colors.Color color1 = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, Convert.ToInt16(color1_string));
                                                                    ltr.Color = color1;
                                                                }

                                                                if (color1_string.ToLower().Contains(",") == true)
                                                                {
                                                                    int idx1 = color1_string.IndexOf(",", 0);
                                                                    byte R = Convert.ToByte(color1_string.Substring(0, idx1));
                                                                    int idx2 = color1_string.IndexOf(",", idx1 + 1);
                                                                    byte G = Convert.ToByte(color1_string.Substring(idx1 + 1, idx2 - idx1 - 1));
                                                                    byte B = Convert.ToByte(color1_string.Substring(idx2 + 1, color1_string.Length - idx2 - 1));
                                                                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(R, G, B);
                                                                }
                                                            }

                                                            //lineweight

                                                            if (dt1.Rows[j][5] != DBNull.Value)
                                                            {
                                                                string val2 = Convert.ToString(dt1.Rows[j][5]);
                                                                if (list_lw.Contains(val2) == true)
                                                                {
                                                                    try
                                                                    {
                                                                        switch (val2)
                                                                        {
                                                                            case "ByLineWeightDefault":
                                                                                ltr.LineWeight = LineWeight.ByLineWeightDefault;
                                                                                break;
                                                                            case "LineWeight000":
                                                                                ltr.LineWeight = LineWeight.LineWeight000;
                                                                                break;
                                                                            case "LineWeight005":
                                                                                ltr.LineWeight = LineWeight.LineWeight005;
                                                                                break;
                                                                            case "LineWeight009":
                                                                                ltr.LineWeight = LineWeight.LineWeight009;
                                                                                break;
                                                                            case "LineWeight013":
                                                                                ltr.LineWeight = LineWeight.LineWeight013;
                                                                                break;
                                                                            case "LineWeight015":
                                                                                ltr.LineWeight = LineWeight.LineWeight015;
                                                                                break;
                                                                            case "LineWeight018":
                                                                                ltr.LineWeight = LineWeight.LineWeight018;
                                                                                break;
                                                                            case "LineWeight020":
                                                                                ltr.LineWeight = LineWeight.LineWeight020;
                                                                                break;
                                                                            case "LineWeight025":
                                                                                ltr.LineWeight = LineWeight.LineWeight025;
                                                                                break;
                                                                            case "LineWeight030":
                                                                                ltr.LineWeight = LineWeight.LineWeight030;
                                                                                break;
                                                                            case "LineWeight035":
                                                                                ltr.LineWeight = LineWeight.LineWeight035;
                                                                                break;
                                                                            case "LineWeight040":
                                                                                ltr.LineWeight = LineWeight.LineWeight040;
                                                                                break;
                                                                            case "LineWeight050":
                                                                                ltr.LineWeight = LineWeight.LineWeight050;
                                                                                break;
                                                                            case "LineWeight060":
                                                                                ltr.LineWeight = LineWeight.LineWeight060;
                                                                                break;
                                                                            case "LineWeight070":
                                                                                ltr.LineWeight = LineWeight.LineWeight070;
                                                                                break;
                                                                            case "LineWeight080":
                                                                                ltr.LineWeight = LineWeight.LineWeight080;
                                                                                break;
                                                                            case "LineWeight090":
                                                                                ltr.LineWeight = LineWeight.LineWeight090;
                                                                                break;
                                                                            case "LineWeight100":
                                                                                ltr.LineWeight = LineWeight.LineWeight100;
                                                                                break;
                                                                            case "LineWeight106":
                                                                                ltr.LineWeight = LineWeight.LineWeight106;
                                                                                break;
                                                                            case "LineWeight120":
                                                                                ltr.LineWeight = LineWeight.LineWeight120;
                                                                                break;
                                                                            case "LineWeight140":
                                                                                ltr.LineWeight = LineWeight.LineWeight140;
                                                                                break;
                                                                            case "LineWeight158":
                                                                                ltr.LineWeight = LineWeight.LineWeight158;
                                                                                break;
                                                                            case "LineWeight200":
                                                                                ltr.LineWeight = LineWeight.LineWeight200;
                                                                                break;
                                                                            case "LineWeight211":
                                                                                ltr.LineWeight = LineWeight.LineWeight211;
                                                                                break;
                                                                            default:
                                                                                ltr.LineWeight = LineWeight.ByLineWeightDefault;
                                                                                MessageBox.Show("The lineweight specified from\r\n" + nume2 + "\r\n" + "on the source drawing is not present in \r\n" + path1 + "\r\n" + " Please reslove manually.");
                                                                                break;
                                                                        }
                                                                    }
                                                                    catch (System.Exception)
                                                                    {
                                                                        MessageBox.Show("The lineweight specified from\r\n" + nume2 + "\r\n" + "on the source drawing is not present in \r\n" + path1 + "\r\n" + " Please reslove manually.");

                                                                    }


                                                                }
                                                            }


                                                            if (dt1.Rows[j]["VP_CENTER_X"] != DBNull.Value && dt1.Rows[j]["VP_CENTER_Y"] != DBNull.Value && dt1.Rows[j]["VP_LAYER_FREEZE"] != DBNull.Value)
                                                            {
                                                                double x1 = Convert.ToDouble(dt1.Rows[j]["VP_CENTER_X"]);
                                                                double y1 = Convert.ToDouble(dt1.Rows[j]["VP_CENTER_Y"]);

                                                                string val_freeze = Convert.ToString(dt1.Rows[j]["VP_LAYER_FREEZE"]).Replace(" ", "");

                                                                bool vpfreeze = false;
                                                                if (val_freeze.ToLower() == "yes" || val_freeze.ToLower() == "true")
                                                                {
                                                                    vpfreeze = true;
                                                                }

                                                                foreach (ObjectId viewportID in viewportCollection)
                                                                {
                                                                    Viewport viewport = Trans2.GetObject(viewportID, OpenMode.ForWrite, false, true) as Viewport;
                                                                    double x2 = viewport.CenterPoint.X;
                                                                    double y2 = viewport.CenterPoint.Y;


                                                                    if (Math.Abs(x1 - x2) < 1 && Math.Abs(y1 - y2) < 1)
                                                                    {
                                                                        foreach (ObjectId id2 in LayerTable2)
                                                                        {
                                                                            LayerTableRecord vpLayer = Trans2.GetObject(id2, OpenMode.ForWrite) as LayerTableRecord;
                                                                            ObjectIdCollection objectIdCollection = new ObjectIdCollection();
                                                                            objectIdCollection.Add(vpLayer.ObjectId);
                                                                            if (vpLayer.Name.ToLower() == nume2.ToLower())
                                                                            {
                                                                                if (vpfreeze == true)
                                                                                {
                                                                                    viewport.FreezeLayersInViewport(objectIdCollection.GetEnumerator());
                                                                                }
                                                                                else
                                                                                {
                                                                                    viewport.ThawLayersInViewport(objectIdCollection.GetEnumerator());
                                                                                }
                                                                            }

                                                                        }
                                                                    }

                                                                }

                                                            }


                                                            j = dt1.Rows.Count;
                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }


                                    Trans2.Commit();
                                }
                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                Database2.SaveAs(path1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                            }
                        }
                        catch (System.Exception ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

        private void load_from_excel_dwg_and_standards_data()
        {
            dt_stnd = new System.Data.DataTable();
            dt_stnd.Columns.Add("Layer Name", typeof(string));
            dt_stnd.Columns.Add("ON/OFF", typeof(string));
            dt_stnd.Columns.Add("THAW/FREEZE", typeof(string));
            dt_stnd.Columns.Add("Color", typeof(string));
            dt_stnd.Columns.Add("Linetype", typeof(string));
            dt_stnd.Columns.Add("Lineweight", typeof(string));
            dt_stnd.Columns.Add("Transparency", typeof(string));
            dt_stnd.Columns.Add("Plot", typeof(string));
            dt_stnd.Columns.Add("VP_CENTER_X", typeof(double));
            dt_stnd.Columns.Add("VP_CENTER_Y", typeof(double));
            dt_stnd.Columns.Add("VP_LAYER_FREEZE", typeof(string));
            dt_stnd.Columns.Add("Stnd", typeof(string));
            dt_stnd.Columns.Add("Config_type", typeof(string));


            dt_dwg_xl = new System.Data.DataTable();
            dt_dwg_xl.Columns.Add("Drawing", typeof(string));
            dt_dwg_xl.Columns.Add("Stnd", typeof(string));

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
            if (stnd_xl_filename != "")
            {
                bool close_wkbook = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;

                foreach (Microsoft.Office.Interop.Excel.Workbook wbk in Excel1.Workbooks)
                {
                    if (wbk.Name == System.IO.Path.GetFileName(stnd_xl_filename))
                    {
                        close_wkbook = false;
                        Workbook1 = wbk;
                    }
                }

                if (close_wkbook == true)
                {
                    Workbook1 = Excel1.Workbooks.Open(stnd_xl_filename);
                }
                bool exista_tab_dwgs = false;
                foreach (Microsoft.Office.Interop.Excel.Worksheet w1 in Workbook1.Worksheets)
                {
                    if (w1.Name.ToUpper() == "DWGS")
                    {
                        exista_tab_dwgs = true;
                    }
                }

                dt_dwg_xl = load_from_excel_list_of_drawings(Workbook1, dt_dwg_xl);
                if (exista_tab_dwgs == true && dt_dwg_xl.Rows.Count > 0)
                {
                    dt_stnd = load_from_Excel_layer_data(Workbook1, dt_stnd);
                }

                try
                {
                    if (close_wkbook == true)
                    {
                        Workbook1.Close();

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
                finally
                {
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
        }


        private System.Data.DataTable load_layers_from_dwg_one_by_one(string dwg1)
        {
            System.Data.DataTable dt3 = new System.Data.DataTable();
            dt3.Columns.Add("Layer Name", typeof(string));
            dt3.Columns.Add("ON/OFF", typeof(string));
            dt3.Columns.Add("THAW/FREEZE", typeof(string));
            dt3.Columns.Add("Color", typeof(string));
            dt3.Columns.Add("Linetype", typeof(string));
            dt3.Columns.Add("Lineweight", typeof(string));
            dt3.Columns.Add("Transparency", typeof(string));
            dt3.Columns.Add("Plot", typeof(string));



            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    if (System.IO.File.Exists(dwg1) == true)
                    {
                        using (Database Database2 = new Database(false, true))
                        {
                            Database2.ReadDwgFile(dwg1, FileOpenMode.OpenForReadAndAllShare, true, "");
                            //System.IO.FileShare.ReadWrite, false, null);
                            Database2.CloseInput(true);
                            Database2.ResolveXrefs(useThreadEngine: true, doNewOnly: false);
                            HostApplicationServices.WorkingDatabase = Database2;
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                            {
                                LayerTable LayerTable2 = Trans2.GetObject(Database2.LayerTableId, OpenMode.ForRead) as LayerTable;
                                foreach (ObjectId id1 in LayerTable2)
                                {
                                    LayerTableRecord ltr = Trans2.GetObject(id1, OpenMode.ForRead) as LayerTableRecord;
                                    if (ltr != null)
                                    {
                                        dt3.Rows.Add();
                                        dt3.Rows[dt3.Rows.Count - 1][0] = ltr.Name;
                                        if (ltr.IsOff == false)
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][1] = "ON";
                                        }
                                        else
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][1] = "OFF";
                                        }

                                        if (ltr.IsFrozen == false)
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][2] = "THAWED";
                                        }
                                        else
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][2] = "FROZEN";
                                        }

                                        Autodesk.AutoCAD.Colors.Color color1 = ltr.Color;
                                        if (color1.IsByAci == true)
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][3] = color1.ColorIndex;
                                        }
                                        else
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][3] = color1.Red + "," + color1.Green + "," + color1.Blue;
                                        }

                                        LinetypeTableRecord linetype1 = Trans2.GetObject(ltr.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;

                                        if (linetype1 != null)
                                        {
                                            dt3.Rows[dt3.Rows.Count - 1][4] = linetype1.Name;
                                        }

                                        dt3.Rows[dt3.Rows.Count - 1][5] = ltr.LineWeight;
                                        dt3.Rows[dt3.Rows.Count - 1][6] = Convert.ToString(ltr.Transparency).Replace("(", "").Replace(")", "");

                                        string yesno = "YES";
                                        if (ltr.IsPlottable == false) yesno = "NO";

                                        dt3.Rows[dt3.Rows.Count - 1][7] = yesno;

                                    }
                                }
                            }
                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                        }
                    }

                }
            }
            return dt3;
        }

        private void button_export_layers_from_selection_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                if (dt_dwg_selected != null && dt_dwg_selected.Rows.Count > 0)
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
                    Excel1.Visible = true;

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Add();

                    List<string> lista_nume = new List<string>();

                    for (int i = 0; i < dt_dwg_selected.Rows.Count; ++i)
                    {
                        if (dt_dwg_selected.Rows[i][0] != DBNull.Value)
                        {
                            string file1 = Convert.ToString(dt_dwg_selected.Rows[i][0]);
                            if (System.IO.File.Exists(file1) == true)
                            {
                                System.Data.DataTable dt1 = load_layers_from_dwg_one_by_one(file1);
                                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets.Add(Before: Workbook1.Worksheets[1]);

                                int maxRows = dt1.Rows.Count;
                                int maxCols = dt1.Columns.Count;
                                char col1 = (char)(64 + maxCols);
                                W1.Range["D:D"].NumberFormat = "@";
                                Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A4:" + col1 + Convert.ToString(4 + maxRows - 1)];
                                range1.ClearContents();
                                object[,] values1 = new object[maxRows, maxCols];
                                for (int k = 0; k < maxRows; ++k)
                                {
                                    for (int j = 0; j < maxCols; ++j)
                                    {
                                        if (dt1.Rows[k][j] != DBNull.Value)
                                        {
                                            values1[k, j] = Convert.ToString(dt1.Rows[k][j]);
                                        }
                                    }
                                }
                                for (int k = 0; k < dt1.Columns.Count; ++k)
                                {
                                    W1.Cells[3, k + 1].value2 = dt1.Columns[k].ColumnName.ToUpper();
                                }
                                range1.Value2 = values1;
                                W1.Columns["A"].ColumnWidth = 50;
                                W1.Columns["B:H"].ColumnWidth = 22;
                                W1.Range["A1"].Value2 = file1;
                                W1.Range["A2"].Value2 = "OVERALL LAYER SETTINGS";

                                string tab1 = System.IO.Path.GetFileNameWithoutExtension(file1);
                                string tab0 = tab1;
                                if (lista_nume.Contains(tab1) == false)
                                {
                                    W1.Name = tab1;
                                    lista_nume.Add(tab1);
                                }
                                else
                                {
                                    int idx = 1;
                                    do
                                    {
                                        tab1 = tab0 + (idx).ToString();
                                        if (lista_nume.Contains(tab1) == false)
                                        {
                                            W1.Name = tab1;
                                            lista_nume.Add(tab1);
                                        }
                                        ++idx;
                                    }
                                    while (lista_nume.Contains(tab1) == true);
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

            set_enable_true();



        }
    }

    public static class att_sync
    {
        public static void SynchronizeAttributes(this BlockTableRecord target)
        {
            if (target == null)
                throw new ArgumentNullException("btr");

            Transaction trans1 = target.Database.TransactionManager.TopTransaction;
            if (trans1 == null)
                throw new AcRx.Exception(ErrorStatus.NoActiveTransactions);


            RXClass attDefClass = RXClass.GetClass(typeof(AttributeDefinition));
            List<AttributeDefinition> attDefs = new List<AttributeDefinition>();
            foreach (ObjectId id in target)
            {
                if (id.ObjectClass == attDefClass)
                {
                    AttributeDefinition attDef = (AttributeDefinition)trans1.GetObject(id, OpenMode.ForRead);
                    attDefs.Add(attDef);
                }
            }
            foreach (ObjectId id in target.GetBlockReferenceIds(true, false))
            {
                BlockReference br = (BlockReference)trans1.GetObject(id, OpenMode.ForWrite);
                br.ResetAttributes(attDefs);
            }
            if (target.IsDynamicBlock)
            {
                foreach (ObjectId id in target.GetAnonymousBlockIds())
                {
                    BlockTableRecord btr = (BlockTableRecord)trans1.GetObject(id, OpenMode.ForRead);
                    foreach (ObjectId brId in btr.GetBlockReferenceIds(true, false))
                    {
                        BlockReference br = (BlockReference)trans1.GetObject(brId, OpenMode.ForWrite);
                        br.ResetAttributes(attDefs);
                    }
                }
            }
        }

        private static void ResetAttributes(this BlockReference br, List<AttributeDefinition> attDefs)
        {
            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = br.Database.TransactionManager;
            Dictionary<string, string> attValues = new Dictionary<string, string>();
            foreach (ObjectId id in br.AttributeCollection)
            {
                if (!id.IsErased)
                {
                    AttributeReference attRef = (AttributeReference)tm.GetObject(id, OpenMode.ForWrite);
                    attValues.Add(attRef.Tag, attRef.TextString);
                    attRef.Erase();
                }
            }
            foreach (AttributeDefinition attDef in attDefs)
            {
                AttributeReference attRef = new AttributeReference();
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                if (attValues.ContainsKey(attDef.Tag))
                {
                    attRef.TextString = attValues[attDef.Tag.ToUpper()];
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tm.AddNewlyCreatedDBObject(attRef, true);
            }
        }


        public static void AttSync(this BlockTableRecord btr, bool directOnly, bool removeSuperfluous, bool setAttDefValues)
        {
            Database db = btr.Database;
            using (WorkingDatabaseSwitcher wdb = new WorkingDatabaseSwitcher(db))
            {
                using (Transaction t = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)t.GetObject(db.BlockTableId, OpenMode.ForRead);

                    IEnumerable<AttributeDefinition> attdefs = btr.Cast<ObjectId>()
                        .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                        .Select(n => (AttributeDefinition)t.GetObject(n, OpenMode.ForRead))
                        .Where(n => !n.Constant);


                    foreach (ObjectId brId in btr.GetBlockReferenceIds(directOnly, false))
                    {
                        BlockReference br = (BlockReference)t.GetObject(brId, OpenMode.ForWrite);

                        if (br.Name != btr.Name)
                            continue;


                        IEnumerable<AttributeReference> attrefs = br.AttributeCollection.Cast<ObjectId>()
                            .Select(n => (AttributeReference)t.GetObject(n, OpenMode.ForWrite));


                        IEnumerable<string> dtags = attdefs.Select(n => n.Tag);

                        IEnumerable<string> rtags = attrefs.Select(n => n.Tag);


                        if (removeSuperfluous)
                            foreach (AttributeReference attref in attrefs.Where(n => rtags
                                .Except(dtags).Contains(n.Tag)))
                                attref.Erase(true);


                        foreach (AttributeReference attref in attrefs.Where(n => dtags
                            .Join(rtags, a => a, b => b, (a, b) => a).Contains(n.Tag)))
                        {
                            AttributeDefinition ad = attdefs.First(n => n.Tag == attref.Tag);


                            string value = attref.TextString;
                            attref.SetAttributeFromBlock(ad, br.BlockTransform);

                            attref.TextString = value;

                            if (attref.IsMTextAttribute)
                            {

                            }


                            if (setAttDefValues)
                                attref.TextString = ad.TextString;

                            attref.AdjustAlignment(db);
                        }


                        IEnumerable<AttributeDefinition> attdefsNew = attdefs.Where(n => dtags
                            .Except(rtags).Contains(n.Tag));

                        foreach (AttributeDefinition ad in attdefsNew)
                        {
                            AttributeReference attref = new AttributeReference();
                            attref.SetAttributeFromBlock(ad, br.BlockTransform);
                            attref.AdjustAlignment(db);
                            br.AttributeCollection.AppendAttribute(attref);
                            t.AddNewlyCreatedDBObject(attref, true);
                        }
                    }
                    btr.UpdateAnonymousBlocks();
                    t.Commit();
                }

                if (btr.IsDynamicBlock)
                {
                    using (Transaction t = db.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId id in btr.GetAnonymousBlockIds())
                        {
                            BlockTableRecord _btr = (BlockTableRecord)t.GetObject(id, OpenMode.ForWrite);


                            IEnumerable<AttributeDefinition> attdefs = btr.Cast<ObjectId>()
                                .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                                 .Select(n => (AttributeDefinition)t.GetObject(n, OpenMode.ForRead));


                            IEnumerable<AttributeDefinition> attdefs2 = _btr.Cast<ObjectId>()
                                 .Where(n => n.ObjectClass.Name == "AcDbAttributeDefinition")
                               .Select(n => (AttributeDefinition)t.GetObject(n, OpenMode.ForWrite));



                            IEnumerable<string> dtags = attdefs.Select(n => n.Tag);
                            IEnumerable<string> dtags2 = attdefs2.Select(n => n.Tag);

                            foreach (AttributeDefinition attdef in attdefs2.Where(n => !dtags.Contains(n.Tag)))
                            {
                                attdef.Erase(true);
                            }


                            foreach (AttributeDefinition attdef in attdefs.Where(n => dtags
                               .Join(dtags2, a => a, b => b, (a, b) => a).Contains(n.Tag)))
                            {
                                AttributeDefinition ad = attdefs2.First(n => n.Tag == attdef.Tag);
                                ad.Position = attdef.Position;
                                ad.TextStyleId = attdef.TextStyleId;

                                if (setAttDefValues)
                                    ad.TextString = attdef.TextString;

                                ad.Tag = attdef.Tag;
                                ad.Prompt = attdef.Prompt;

                                ad.LayerId = attdef.LayerId;
                                ad.Rotation = attdef.Rotation;
                                ad.LinetypeId = attdef.LinetypeId;
                                ad.LineWeight = attdef.LineWeight;
                                ad.LinetypeScale = attdef.LinetypeScale;
                                ad.Annotative = attdef.Annotative;
                                ad.Color = attdef.Color;
                                ad.Height = attdef.Height;
                                ad.HorizontalMode = attdef.HorizontalMode;
                                ad.Invisible = attdef.Invisible;
                                ad.IsMirroredInX = attdef.IsMirroredInX;
                                ad.IsMirroredInY = attdef.IsMirroredInY;
                                ad.Justify = attdef.Justify;
                                ad.LockPositionInBlock = attdef.LockPositionInBlock;
                                ad.MaterialId = attdef.MaterialId;
                                ad.Oblique = attdef.Oblique;
                                ad.Thickness = attdef.Thickness;
                                ad.Transparency = attdef.Transparency;
                                ad.VerticalMode = attdef.VerticalMode;
                                ad.Visible = attdef.Visible;
                                ad.WidthFactor = attdef.WidthFactor;

                                ad.CastShadows = attdef.CastShadows;
                                ad.Constant = attdef.Constant;
                                ad.FieldLength = attdef.FieldLength;
                                ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible;
                                ad.Preset = attdef.Preset;
                                ad.Prompt = attdef.Prompt;
                                ad.Verifiable = attdef.Verifiable;

                                ad.AdjustAlignment(db);
                            }


                            foreach (AttributeDefinition attdef in attdefs.Where(n => !dtags2.Contains(n.Tag)))
                            {
                                AttributeDefinition ad = new AttributeDefinition();
                                ad.SetDatabaseDefaults();
                                ad.Position = attdef.Position;
                                ad.TextStyleId = attdef.TextStyleId;
                                ad.TextString = attdef.TextString;
                                ad.Tag = attdef.Tag;
                                ad.Prompt = attdef.Prompt;

                                ad.LayerId = attdef.LayerId;
                                ad.Rotation = attdef.Rotation;
                                ad.LinetypeId = attdef.LinetypeId;
                                ad.LineWeight = attdef.LineWeight;
                                ad.LinetypeScale = attdef.LinetypeScale;
                                ad.Annotative = attdef.Annotative;
                                ad.Color = attdef.Color;
                                ad.Height = attdef.Height;
                                ad.HorizontalMode = attdef.HorizontalMode;
                                ad.Invisible = attdef.Invisible;
                                ad.IsMirroredInX = attdef.IsMirroredInX;
                                ad.IsMirroredInY = attdef.IsMirroredInY;
                                ad.Justify = attdef.Justify;
                                ad.LockPositionInBlock = attdef.LockPositionInBlock;
                                ad.MaterialId = attdef.MaterialId;
                                ad.Oblique = attdef.Oblique;
                                ad.Thickness = attdef.Thickness;
                                ad.Transparency = attdef.Transparency;
                                ad.VerticalMode = attdef.VerticalMode;
                                ad.Visible = attdef.Visible;
                                ad.WidthFactor = attdef.WidthFactor;

                                ad.CastShadows = attdef.CastShadows;
                                ad.Constant = attdef.Constant;
                                ad.FieldLength = attdef.FieldLength;
                                ad.ForceAnnoAllVisible = attdef.ForceAnnoAllVisible;
                                ad.Preset = attdef.Preset;
                                ad.Prompt = attdef.Prompt;
                                ad.Verifiable = attdef.Verifiable;

                                _btr.AppendEntity(ad);
                                t.AddNewlyCreatedDBObject(ad, true);
                                ad.AdjustAlignment(db);
                            }

                            _btr.AttSync(directOnly, removeSuperfluous, setAttDefValues);
                        }

                        btr.UpdateAnonymousBlocks();
                        t.Commit();
                    }
                }
            }
        }
    }

    public sealed class WorkingDatabaseSwitcher : IDisposable
    {
        private Database prevDb = null;


        public WorkingDatabaseSwitcher(Database db)
        {
            prevDb = HostApplicationServices.WorkingDatabase;
            HostApplicationServices.WorkingDatabase = db;
        }

        /// Возвращаем свойству <c>HostApplicationServices.WorkingDatabase</c> прежнее значение
        /// </summary>
        public void Dispose()
        {
            HostApplicationServices.WorkingDatabase = prevDb;
        }
    }

}
