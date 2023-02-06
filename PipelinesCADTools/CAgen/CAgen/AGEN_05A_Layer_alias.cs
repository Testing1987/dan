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

namespace Alignment_mdi
{
    public partial class AGEN_Layer_alias : Form
    {
        System.Data.DataTable dt_layers = null;
        bool is_saved = false;
        List<string> lista_layere = null;

        public AGEN_Layer_alias()
        {
            InitializeComponent();
        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_layer_alias);
            lista_butoane.Add(button_new_alias_table);
            lista_butoane.Add(button_open_excel_layer_alias);
            lista_butoane.Add(button_scan_layers);
            lista_butoane.Add(button_show_scan_and_draw_crossings);
            lista_butoane.Add(button_transfer_selected);
            lista_butoane.Add(button_transfer_all);
            lista_butoane.Add(button_transfer_to_excel_layer_alias);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_load_layer_alias);
            lista_butoane.Add(button_new_alias_table);
            lista_butoane.Add(button_open_excel_layer_alias);
            lista_butoane.Add(button_scan_layers);
            lista_butoane.Add(button_show_scan_and_draw_crossings);
            lista_butoane.Add(button_transfer_selected);
            lista_butoane.Add(button_transfer_all);
            lista_butoane.Add(button_transfer_to_excel_layer_alias);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        private void button_load_layer_alias_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
                if (Ag != null)
                {
                    if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
                    {
                        string ProjFolder_for_layer_alias = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();
                        string fisier_alias = ProjFolder_for_layer_alias + _AGEN_mainform.layer_alias_excel_name;
                        if (System.IO.File.Exists(fisier_alias) == true)
                        {
                            _AGEN_mainform.dt_layer_alias = Load_existing_layer_alias_from_excel(fisier_alias);
                            if (_AGEN_mainform.dt_layer_alias.Rows.Count > 0)
                            {
                                label_layer_alias.Text = "Layer Alias loaded";
                                label_layer_alias.ForeColor = Color.LimeGreen;
                                is_saved = true;
                                lista_layere = new List<string>();
                                for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.dt_layer_alias.Rows[i][0] != DBNull.Value)
                                    {
                                        lista_layere.Add(Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i][0]));
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("you do not have data in the layer alias file", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                            populate_grid_with_layer_alias_data();
                        }
                        else
                        {
                            _AGEN_mainform.dt_layer_alias = null;
                            lista_layere = null;
                            MessageBox.Show("you do not have the layer alias file", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        _AGEN_mainform.dt_layer_alias = null;
                        lista_layere = null;
                        MessageBox.Show("no project folder found", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }
        public System.Data.DataTable Load_existing_layer_alias_from_excel(string File1)
        {
            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the layer alias data file does not exist");
                return null;
            }
            System.Data.DataTable dt2 = new System.Data.DataTable();
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

                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                try
                {
                    dt2 = Functions.Build_Data_table_layer_alias_from_excel(W1, _AGEN_mainform.Start_row_layer_alias + 1);
                    Workbook1.Close();
                    if (Excel1.Workbooks.Count==0)
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
            return dt2;
        }

        private void populate_grid_with_layer_alias_data()
        {
            dataGridView_layer_alias.DataSource = _AGEN_mainform.dt_layer_alias;
            dataGridView_layer_alias.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_layer_alias.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_layer_alias.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_layer_alias.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_layer_alias.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_layer_alias.EnableHeadersVisualStyles = false;
        }





        private void button_show_scan_and_draw_crossings_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_layer_alias.Hide();
            _AGEN_mainform.tpage_crossing_scan.Show();

        }

        public void Set_layer_alias_label_to_green()
        {
            label_layer_alias.Text = "Layer Alias loaded";
            label_layer_alias.ForeColor = Color.LimeGreen;
        }

        public void Set_layer_alias_label_to_red()
        {
            label_layer_alias.Text = "No Layer Alias Loaded";
            label_layer_alias.ForeColor = Color.Red;
        }
        private void button_transfer_to_excel_layer_alias_Click(object sender, EventArgs e)
        {
            

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.layer_alias_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.layer_alias_excel_name + " file");
                return;
            }

            string ProjF_layer_alias = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();
            if (ProjF_layer_alias.Substring(ProjF_layer_alias.Length - 1, 1) != "\\")
            {
                ProjF_layer_alias = ProjF_layer_alias + "\\";
            }
            if (System.IO.Directory.Exists(ProjF_layer_alias) == true && System.IO.File.Exists(ProjF_layer_alias + _AGEN_mainform.layer_alias_excel_name) == true)
            {

            }
            else
            {
                MessageBox.Show("no layer alias specified", "agen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }




            try
            {
                set_enable_false();
                if (_AGEN_mainform.dt_layer_alias != null)
                {
                    if (_AGEN_mainform.dt_layer_alias.Rows.Count > 0)
                    {
                        Save_data_inside_layer_alias(_AGEN_mainform.dt_layer_alias, ProjF_layer_alias + _AGEN_mainform.layer_alias_excel_name);
                        is_saved = true;
                        Set_layer_alias_label_to_green();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }


        private void Save_data_inside_layer_alias(System.Data.DataTable dt_alias, string path1)
        {
            if (dt_alias != null && dt_alias.Rows.Count > 0)
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
                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                try
                {
                    Workbook1 = Excel1.Workbooks.Open(path1);
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    W1.Name = "alias";
                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_layer_alias, _AGEN_mainform.Start_row_layer_alias, "@");
                    Functions.Create_header_layer_alias_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), null);

                    Workbook1.Save();
                    Workbook1.Close();

                    if (Excel1.Workbooks.Count==0)
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

                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
        }


        private void button_open_excel_layer_alias_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                string ProjFolder_for_layer_alias = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();
                if (System.IO.Directory.Exists(ProjFolder_for_layer_alias) == true)
                {
                    string fisier_alias = ProjFolder_for_layer_alias + _AGEN_mainform.layer_alias_excel_name;
                    if (System.IO.File.Exists(fisier_alias) == false)
                    {
                        set_enable_true();

                        MessageBox.Show("the layer alias data file does not exist");
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
                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_alias);
                }
                else
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    MessageBox.Show("the project folder does not exist");
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            set_enable_true();

        }

        private void button_new_alias_table_Click(object sender, EventArgs e)
        {
            try
            {
                set_enable_false();
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder_without_segment();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_la = ProjFolder + _AGEN_mainform.layer_alias_excel_name;
                    _AGEN_mainform.dt_layer_alias = null;
                    if (System.IO.File.Exists(fisier_la) == false)
                    {
                        _AGEN_mainform.dt_layer_alias = Functions.Creaza_layer_alias_datatable_structure();
                        creaza_new_layer_alias_file(_AGEN_mainform.dt_layer_alias, fisier_la);
                        populate_grid_with_layer_alias_data();
                        lista_layere = new List<string>();
                    }
                    else
                    {
                        MessageBox.Show("layer alias exists!\r\nOperation aborted");
                        _AGEN_mainform.dt_layer_alias = null;
                        lista_layere = null;
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!\r\nnoperation aborted");
                    _AGEN_mainform.dt_layer_alias = null;
                    lista_layere = null;


                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.dt_layer_alias = null;
                lista_layere = null;

                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void creaza_new_layer_alias_file(System.Data.DataTable dt_la, string fis_la)
        {

            if (dt_la != null)
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

                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;


                try
                {
                    if (dt_la != null)
                    {
                        Workbook1 = Excel1.Workbooks.Add();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                        Functions.Create_header_layer_alias_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), dt_la);
                        Workbook1.SaveAs(fis_la);
                        Workbook1.Close();
                    }

                    if (Excel1.Workbooks.Count==0)
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

                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }




        }

        private void button_scan_layers_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                dt_layers = new System.Data.DataTable();
                dt_layers.Columns.Add("DWG Layer", typeof(string));
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        if (lista_layere == null) lista_layere = new List<string>();
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        foreach (ObjectId id1 in LayerTable1)
                        {
                            LayerTableRecord ltr = Trans1.GetObject(id1, OpenMode.ForRead) as LayerTableRecord;
                            if (ltr != null)
                            {
                                string nume1 = ltr.Name;
                                if (nume1 != "0" && nume1 != "Defpoints" && nume1.Contains("|") == false && nume1.Contains("*") == false)
                                {
                                    if (lista_layere.Contains(nume1) == false)
                                    {
                                        dt_layers.Rows.Add();
                                        dt_layers.Rows[dt_layers.Rows.Count - 1][0] = nume1;
                                    }

                                }

                            }
                        }


                    }
                }
                if (dt_layers.Rows.Count > 0)
                {
                    populate_grid_with_layer_dwg_data();
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

        private void populate_grid_with_layer_dwg_data()
        {

            dataGridView_dwg_layers.DataSource = dt_layers;
            dataGridView_dwg_layers.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_dwg_layers.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_dwg_layers.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_dwg_layers.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_dwg_layers.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_dwg_layers.EnableHeadersVisualStyles = false;

        }

        private void button_transfer_selected_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewSelectedCellCollection col1 = dataGridView_dwg_layers.SelectedCells;
                if (col1.Count > 0)
                {
                    foreach (DataGridViewCell cell1 in col1)
                    {
                        string ln = Convert.ToString(cell1.Value);
                        if (ln != null && ln != "" && lista_layere.Contains(ln) == false)
                        {
                            _AGEN_mainform.dt_layer_alias.Rows.Add();
                            _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][0] = ln;
                            _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][2] = "NO";
                            _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][20] = "YES";

                            lista_layere.Add(ln);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button_transfer_all_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_layers != null && dt_layers.Rows.Count > 0)
                {

                    for (int i = 0; i < dt_layers.Rows.Count; ++i)
                    {
                        if (dt_layers.Rows[i][0] != DBNull.Value)
                        {
                            string ln = Convert.ToString(dt_layers.Rows[i][0]);
                            if (lista_layere.Contains(ln) == false)
                            {
                                _AGEN_mainform.dt_layer_alias.Rows.Add();
                                _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][0] = ln;
                                _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][2] = "NO";
                                _AGEN_mainform.dt_layer_alias.Rows[_AGEN_mainform.dt_layer_alias.Rows.Count - 1][20] = "YES";
                                lista_layere.Add(ln);
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

        public void update_alias_file_from_crossing(System.Data.DataTable dt_alias, string path1)
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

                if (Excel1.Workbooks.Count==0) Excel1.Visible = _AGEN_mainform.ExcelVisible;



                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                if (System.IO.File.Exists(path1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();

                }

                else
                {
                    Workbook1 = Excel1.Workbooks.Open(path1);
                }

                if (_AGEN_mainform.dt_layer_alias != null)
                {
                    if (_AGEN_mainform.dt_layer_alias.Rows.Count > 0)
                    {
                        for (int i = 0; i < _AGEN_mainform.dt_layer_alias.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_layer_alias.Rows[i][20] != DBNull.Value)
                            {
                                string display = Convert.ToString(_AGEN_mainform.dt_layer_alias.Rows[i][20]);
                                if (display.ToUpper() == "YES")
                                {
                                    _AGEN_mainform.dt_layer_alias.Rows[i][20] = "YES";
                                }
                                if (display.ToUpper() == "NO")
                                {
                                    _AGEN_mainform.dt_layer_alias.Rows[i][20] = "NO";
                                }
                                else
                                {
                                    _AGEN_mainform.dt_layer_alias.Rows[i][20] = "YES";
                                }
                            }
                            else
                            {
                                _AGEN_mainform.dt_layer_alias.Rows[i][20] = "YES";
                            }
                        }
                    }
                }

                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                W1.Name = "alias";
                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                if (segment1 == "not defined") segment1 = "";
                Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.dt_layer_alias, _AGEN_mainform.Start_row_layer_alias, "@");
                Functions.Create_header_layer_alias_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), null);

                try
                {
                    if (System.IO.File.Exists(path1) == false)
                    {
                        Workbook1.SaveAs(path1);
                    }
                    else
                    {
                        Workbook1.Save();
                    }

                    Workbook1.Close();
                    if (Excel1.Workbooks.Count==0)
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
}
