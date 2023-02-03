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
using Microsoft.Office.Interop.Excel;

namespace Alignment_mdi
{
    public partial class AGEN_Sheet_Generation : Form
    {
        bool Freeze_operations = false;
        private ContextMenuStrip ContextMenuStrip_open_alignment;
        System.Data.DataTable dt_image = null;

        System.Data.DataTable Display_dt;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_cut_sheets);
            lista_butoane.Add(button_output_location);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_cut_sheets);

            lista_butoane.Add(button_output_location);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public AGEN_Sheet_Generation()
        {
            InitializeComponent();
            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Open selected drawing" };
            toolStripMenuItem1.Click += open_alignment_Click;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Add drawings" };
            toolStripMenuItem2.Click += add_dwgs_Click;


            var toolStripMenuItem3 = new ToolStripMenuItem { Text = "Remove drawing" };
            toolStripMenuItem3.Click += remove_selected_dwg_Click;

            var toolStripMenuItem4 = new ToolStripMenuItem { Text = "Clear drawing list" };
            toolStripMenuItem4.Click += remove_all_dwg_Click;

            ContextMenuStrip_open_alignment = new ContextMenuStrip();
            ContextMenuStrip_open_alignment.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2, toolStripMenuItem3, toolStripMenuItem4 });


        }

        private void add_dwgs_Click(object sender, EventArgs e)
        {
            if (Display_dt == null)
            {
                Display_dt = Functions.Creaza_display_datatable_structure();
            }
            else if (Display_dt.Rows.Count == 0)
            {
                Display_dt = Functions.Creaza_display_datatable_structure();
            }
            string Col_dwg_name = "DwgNo";
            string Col_M1 = "StaBeg";
            string Col_M2 = "StaEnd";


            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "alignment sheet (*.dwg)|*.dwg";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    for (int i = 0; i < fbd.FileNames.Count(); ++i)
                    {
                        string File1 = fbd.FileNames[i];
                        bool Add1 = true;

                        if (Display_dt.Rows.Count > 0)
                        {
                            for (int k = 0; k < Display_dt.Rows.Count; ++k)
                            {
                                if (Display_dt.Rows[k][Col_dwg_name].ToString() == File1)
                                {
                                    Add1 = false;
                                    k = Display_dt.Rows.Count;
                                }
                            }
                        }

                        if (Add1 == true)
                        {
                            Display_dt.Rows.Add();
                            Display_dt.Rows[Display_dt.Rows.Count - 1][Col_dwg_name] = File1;
                            string nume1 = System.IO.Path.GetFileNameWithoutExtension(File1);

                            if (_AGEN_mainform.dt_sheet_index != null)
                            {
                                if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                                {
                                    for (int j = 0; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                    {
                                        if (_AGEN_mainform.dt_sheet_index.Rows[j][Col_dwg_name] != DBNull.Value)
                                        {
                                            if (nume1 == _AGEN_mainform.dt_sheet_index.Rows[j][Col_dwg_name].ToString() &&
                                                                                 _AGEN_mainform.dt_sheet_index.Rows[j][Col_M1] != DBNull.Value &&
                                                                                                _AGEN_mainform.dt_sheet_index.Rows[j][Col_M2] != DBNull.Value)
                                            {
                                                Display_dt.Rows[Display_dt.Rows.Count - 1][Col_M1] = _AGEN_mainform.dt_sheet_index.Rows[j][Col_M1];
                                                Display_dt.Rows[Display_dt.Rows.Count - 1][Col_M2] = _AGEN_mainform.dt_sheet_index.Rows[j][Col_M2];
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            Display_dt = Functions.Sort_data_table(Display_dt, Col_M1);

            if (Display_dt.Rows.Count > 0)
            {
                dataGridView_align_created.DataSource = Display_dt;
                dataGridView_align_created.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_align_created.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_align_created.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_align_created.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_align_created.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_align_created.EnableHeadersVisualStyles = false;
            }

        }


        private void remove_selected_dwg_Click(object sender, EventArgs e)
        {
            if (dataGridView_align_created.RowCount > 0)
            {
                int Index1 = dataGridView_align_created.CurrentCell.RowIndex;
                if (Index1 == -1)
                {
                    return;
                }

                dataGridView_align_created.Rows.RemoveAt(Index1);
            }
        }

        private void remove_all_dwg_Click(object sender, EventArgs e)
        {
            if (Display_dt != null)
            {
                Display_dt = Functions.Creaza_display_datatable_structure();
            }

            dataGridView_align_created.DataSource = "";
        }

        private void open_alignment_Click(object sender, EventArgs e)
        {
            try
            {

                int Index1 = dataGridView_align_created.CurrentCell.RowIndex;
                if (Display_dt != null)
                {
                    if (Display_dt.Rows.Count - 1 >= Index1)
                    {
                        string fisier_generat = Display_dt.Rows[Index1][_AGEN_mainform.Col_dwg_name].ToString();
                        if (System.IO.File.Exists(fisier_generat) == true)
                        {

                            bool is_opened = false;
                            DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                            foreach (Document Doc in DocumentManager1)
                            {
                                if (Doc.Name == fisier_generat)
                                {
                                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Doc;
                                    is_opened = true;

                                }

                            }

                            if (is_opened == false)
                            {
                                DocumentCollectionExtension.Open(DocumentManager1, fisier_generat, false);
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






        public void hide_prof_options_at_generation()
        {
            panel_dan.Visible = false;
        }

        public void set_checkBox_gen_plan_view_off()
        {
            checkBox_plan_view.Checked = false;
        }
        public void set_checkBox_gen_plan_view_on()
        {
            checkBox_plan_view.Checked = true;
        }
        public void set_checkBox_gen_profile_off()
        {
            checkBox_profile.Checked = false;
        }
        public void set_checkBox_gen_profile_on()
        {
            checkBox_profile.Checked = true;
        }


        public void set_checkBox_gen_ownership_off()
        {
            checkBox_ownership.Checked = false;
        }
        public void set_checkBox_gen_ownership_on()
        {
            checkBox_ownership.Checked = true;
        }

        public void set_checkBox_gen_crossing_off()
        {
            checkBox_crossing.Checked = false;
        }
        public void set_checkBox_gen_crossing_on()
        {
            checkBox_crossing.Checked = true;
        }
        public void set_checkBox_gen_materials_off()
        {
            checkBox_materials.Checked = false;
        }
        public void set_checkBox_gen_materials_on()
        {
            checkBox_materials.Checked = true;
        }

        private void draw_custom_vp_rectangles(Autodesk.AutoCAD.DatabaseServices.Transaction Trans1, BlockTableRecord Btrecord, int lr, Point3d ms_point, double width1, double height1, string file_name)
        {
            if (checkBox_draw_rectangle_custom.Checked == true)
            {
                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);

                Polyline polyVP1 = new Polyline();

                polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                polyVP1.ColorIndex = 1;

                polyVP1.AddVertexAt(0, new Point2d(ms_point.X - lr * (width1 / 2) / 1, ms_point.Y - (height1 / 2) / 1), 0, 0, 0);
                polyVP1.AddVertexAt(1, new Point2d(ms_point.X - lr * (width1 / 2) / 1, ms_point.Y + (height1 / 2) / 1), 0, 0, 0);
                polyVP1.AddVertexAt(2, new Point2d(ms_point.X + lr * (width1 / 2) / 1, ms_point.Y + (height1 / 2) / 1), 0, 0, 0);
                polyVP1.AddVertexAt(3, new Point2d(ms_point.X + lr * (width1 / 2) / 1, ms_point.Y - (height1 / 2) / 1), 0, 0, 0);
                polyVP1.Closed = true;

                Btrecord.AppendEntity(polyVP1);
                Trans1.AddNewlyCreatedDBObject(polyVP1, true);

                MText mtext1 = new MText();
                mtext1.Contents = file_name;
                mtext1.TextHeight = 0.75;
                mtext1.Location = new Point3d(ms_point.X - lr * (width1 / 2) / 1, 0.5 + ms_point.Y + (height1 / 2) / 1, 0);
                mtext1.Attachment = AttachmentPoint.BottomLeft;
                mtext1.Rotation = 0;
                mtext1.Layer = _AGEN_mainform.layer_no_plot;

                Btrecord.AppendEntity(mtext1);
                Trans1.AddNewlyCreatedDBObject(mtext1, true);
            }
        }

        private void button_cut_sheets_Click(object sender, EventArgs e)
        {
           

            double h_nd = 0;
            double w_nd = 0;
            double x_nd_ms = 0;
            double y_nd_ms = 0;
            double x_nd_ps = 0;
            double y_nd_ps = 0;
            double sep_nd = 0;
            double scale_nd = 0;


            DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (DocumentManager1.Count == 0)
            {
                string strTemplatePath = "acad.dwt";
                Document acDoc = DocumentManager1.Add(strTemplatePath);
                DocumentManager1.MdiActiveDocument = acDoc;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            if (Ag != null)
            {
                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }

            if (_AGEN_mainform.Exista_viewport_main == false && checkBox_plan_view.Checked == true)
            {
                MessageBox.Show("no main viewport dimension specified\r\nOperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();

                _AGEN_mainform.tpage_viewport_settings.Show();
                return;
            }

            if (_AGEN_mainform.Exista_viewport_cross == false && checkBox_crossing.Checked == true)
            {
                MessageBox.Show("no crossing viewport dimension specified\r\nOperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();

                _AGEN_mainform.tpage_viewport_settings.Show();
                return;
            }

            if (_AGEN_mainform.Exista_viewport_owner == false && checkBox_ownership.Checked == true)
            {
                MessageBox.Show("no ownership viewport dimension specified\r\nOperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();

                _AGEN_mainform.tpage_viewport_settings.Show();
                return;
            }

            if (_AGEN_mainform.Exista_viewport_prof == false && checkBox_profile.Checked == true)
            {
                MessageBox.Show("no profile viewports dimension specified\r\nOperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();

                _AGEN_mainform.tpage_viewport_settings.Show();
                return;
            }

            if (_AGEN_mainform.Exista_viewport_mat == false && checkBox_materials.Checked == true)
            {
                MessageBox.Show("no materials viewports dimension specified\r\nOperation aborted");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();
                _AGEN_mainform.tpage_tblk_attrib.Hide();
                _AGEN_mainform.tpage_sheetindex.Hide();
                _AGEN_mainform.tpage_layer_alias.Hide();
                _AGEN_mainform.tpage_crossing_scan.Hide();
                _AGEN_mainform.tpage_crossing_draw.Hide();
                _AGEN_mainform.tpage_profilescan.Hide();
                _AGEN_mainform.tpage_profdraw.Hide();
                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();

                _AGEN_mainform.tpage_viewport_settings.Show();
                return;

            }

            if (_AGEN_mainform.dt_sheet_index == null)
            {
                MessageBox.Show("no sheet index data found\r\nOperation aborted");
                return;
            }

            if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("no sheet index data found\r\nOperation aborted");
                return;
            }
            try
            {
                try
                {
                    if (Freeze_operations == false)
                    {
                        _AGEN_mainform.tpage_processing.Show();
                        //Ag.WindowState = FormWindowState.Minimized;
                        Freeze_operations = true;
                        _AGEN_mainform.dt_station_equation = null;
                        System.Data.DataTable Dtp = null;
                        if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                        {
                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                            {
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_prof)
                                {
                                    _AGEN_mainform.Vw_prof_height = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                }
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_prop)
                                {
                                    _AGEN_mainform.Vw_prop_height = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                }
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_cross)
                                {
                                    _AGEN_mainform.Vw_cross_height = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                }
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_mat)
                                {
                                    _AGEN_mainform.Vw_mat_height = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                }
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_profband)
                                {
                                    _AGEN_mainform.Vw_profband_height = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                }
                                if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == _AGEN_mainform.nume_banda_no_data)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i][1] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][11] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][12] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][13] != DBNull.Value &&
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i][14] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][15] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][16] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i][17] != DBNull.Value)
                                    {
                                        h_nd = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                        w_nd = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"]);
                                        x_nd_ms = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"]);
                                        y_nd_ms = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"]);
                                        x_nd_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"]);
                                        y_nd_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"]);
                                        sep_nd = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"]);
                                        scale_nd = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"]);
                                    }


                                }
                            }
                        }

                        string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                        {
                            ProjFolder = ProjFolder + "\\";
                        }

                        if (System.IO.Directory.Exists(ProjFolder) == true)
                        {
                            if (checkBox_profile.Checked == true)
                            {
                                string fisier_prof = ProjFolder + _AGEN_mainform.prof_excel_name;

                                if (System.IO.File.Exists(fisier_prof) == true)
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
                                    try
                                    {
                                        Dtp = Load_existing_profile_graph(fisier_prof);
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
                                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("the profile data file does not exist");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    Freeze_operations = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            MessageBox.Show("there is no such a project folder");
                            Freeze_operations = false;
                            return;
                        }

                        if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            MessageBox.Show("sheet index table is empty\r\nOperation aborted");
                            Freeze_operations = false;
                            return;
                        }


                        string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();

                        if (Output_folder.Substring(Output_folder.Length - 1, 1) != "\\")
                        {
                            Output_folder = Output_folder + "\\";
                        }

                        Point3d ms_point = new Point3d();
                        Point3d ps_point_plan_view = new Point3d(_AGEN_mainform.Vw_ps_x, _AGEN_mainform.Vw_ps_y, 0);

                        bool Creaza_new_file = true;

                        if (Display_dt == null)
                        {
                            Creaza_new_file = true;
                        }

                        if (Display_dt != null)
                        {
                            if (Display_dt.Rows.Count == 0)
                            {
                                Creaza_new_file = true;
                            }
                            else
                            {
                                Creaza_new_file = false;
                            }
                        }

                        int lr = 1;
                        if (_AGEN_mainform.Left_to_Right == false) lr = -1;
                        double scale_cust = 1;
                        if (checkBox_custom_vp_scale.Checked == true) scale_cust = _AGEN_mainform.Vw_scale;

                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord Btrecord = Functions.get_modelspace(Trans1, ThisDrawing.Database);
                                if ((checkBox_profile_band.Checked == true || checkBox_draw_rectangle_custom.Checked == true || checkBox_mult_vp_prof.Checked == true || checkBox_no_data_band.Checked == true) && checkBox_delete_vp.Checked == false)
                                {
                                    Btrecord.UpgradeOpen();
                                }

                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {

                                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                        Polyline3d poly3d = null;
                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                        }


                                        if (_AGEN_mainform.dt_station_equation.Columns.Contains("measured") == false)
                                        {
                                            _AGEN_mainform.dt_station_equation.Columns.Add("measured", typeof(double));
                                        }

                                        for (int i = 0; i < _AGEN_mainform.dt_station_equation.Rows.Count; ++i)
                                        {
                                            if (_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"] != DBNull.Value)
                                            {
                                                double x = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End X"]);
                                                double y = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[i]["Reroute End Y"]);
                                                Point3d pt_on_2d = poly2d.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                                double eq_meas = poly2d.GetDistAtPoint(pt_on_2d);
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    double param1 = poly2d.GetParameterAtPoint(pt_on_2d);
                                                    eq_meas = poly3d.GetDistanceAtParameter(param1);
                                                }

                                                _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                            }
                                        }

                                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                                    }
                                }
                                else
                                {
                                    _AGEN_mainform.dt_station_equation = null;
                                }

                                #region creaza new file

                                if (Creaza_new_file == true)
                                {

                                    List<int> lista_generation = new List<int>();

                                    if (comboBox_start.Text != "" & comboBox_end.Text != "")
                                    {
                                        lista_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg(comboBox_start.Text, comboBox_end.Text);
                                    }



                                    if (lista_generation.Count == 0)
                                    {
                                        for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                        {
                                            lista_generation.Add(i);
                                        }
                                    }

                                    bool odd_even = false;
                                    string Template_file_name = "";

                                    if (System.IO.File.Exists(_AGEN_mainform.template2))
                                    {
                                        Template_file_name = _AGEN_mainform.template2;

                                    }

                                    if (System.IO.File.Exists(_AGEN_mainform.template1))
                                    {
                                        Template_file_name = _AGEN_mainform.template1;
                                    }



                                    if (System.IO.File.Exists(_AGEN_mainform.template1) == true && System.IO.File.Exists(_AGEN_mainform.template2) == true && _AGEN_mainform.template1 != _AGEN_mainform.template2)
                                    {
                                        odd_even = true;
                                    }




                                    Document New_doc = DocumentCollectionExtension.Add(DocumentManager1, Template_file_name);
                                    DocumentManager1.MdiActiveDocument = New_doc;

                                    string fname0 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[0]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";

                                    if (System.IO.File.Exists(fname0) == true)
                                    {
                                        try
                                        {
                                            string existingFolder = System.IO.Path.GetDirectoryName(fname0);
                                            existingFolder = existingFolder.EndsWith("\\") ? existingFolder : existingFolder + "\\";

                                            string segmentName = _AGEN_mainform.current_segment;
                                            string date = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
                                            string folderName = $"{segmentName}-{date} by {Environment.UserName}";
                                            string newFolder = existingFolder + folderName;
                                            string destinationPath = System.IO.Path.Combine(newFolder, System.IO.Path.GetFileName(fname0));

                                            if (!System.IO.Directory.Exists(newFolder))
                                            {
                                                System.IO.Directory.CreateDirectory(newFolder);
                                            }
                                            System.IO.File.Move(fname0, destinationPath);
                                        }
                                        catch (System.Exception)
                                        {
                                            MessageBox.Show(fname0 + " cannot be moved.\r\noperation aborted");
                                            Display_dt = Functions.Creaza_display_datatable_structure();
                                            Freeze_operations = false;
                                            _AGEN_mainform.tpage_processing.Hide();
                                            return;

                                        }



                                    }

                                    using (DocumentLock lock2 = New_doc.LockDocument())
                                    {

                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = New_doc.Database.TransactionManager.StartTransaction())
                                        {
                                            BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, New_doc.Database);
                                            BtrecordPS.UpgradeOpen();
                                            Layout Layout1 = Functions.get_first_layout(Trans2, New_doc.Database);
                                            Layout1.UpgradeOpen();
                                            Layout1.LayoutName = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[0]][_AGEN_mainform.Col_dwg_name].ToString();
                                            Trans2.Commit();
                                            New_doc.Database.SaveAs(fname0, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                        }
                                    }



                                    New_doc.CloseAndDiscard();

                                    if (odd_even == true)
                                    {
                                        if (lista_generation.Count > 1)
                                        {
                                            Document New_doc2 = DocumentCollectionExtension.Add(DocumentManager1, _AGEN_mainform.template2);
                                            DocumentManager1.MdiActiveDocument = New_doc2;

                                            string fname2 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[1]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";

                                            if (System.IO.File.Exists(fname2) == true)
                                            {
                                                MessageBox.Show(fname2 + " already exists.\r\noperation aborted");
                                                Display_dt = Functions.Creaza_display_datatable_structure();
                                                Freeze_operations = false;
                                                _AGEN_mainform.tpage_processing.Hide();
                                                return;
                                            }

                                            using (DocumentLock lock2 = New_doc2.LockDocument())
                                            {

                                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = New_doc2.Database.TransactionManager.StartTransaction())
                                                {
                                                    BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, New_doc2.Database);
                                                    BtrecordPS.UpgradeOpen();
                                                    Layout Layout1 = Functions.get_first_layout(Trans2, New_doc2.Database);
                                                    Layout1.UpgradeOpen();
                                                    Layout1.LayoutName = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[1]][_AGEN_mainform.Col_dwg_name].ToString();
                                                    Trans2.Commit();
                                                    New_doc2.Database.SaveAs(fname2, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                }
                                            }



                                            New_doc2.CloseAndDiscard();
                                        }
                                    }


                                    Display_dt = Functions.Creaza_display_datatable_structure();

                                    bool is1 = true;
                                    if (_AGEN_mainform.dt_sheet_index.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_generation.Count; ++i)
                                        {
                                            string Fisier2 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";

                                            if (odd_even == false)
                                            {
                                                if (i > 0)
                                                {
                                                    string Fisier1 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i - 1]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";
                                                    System.IO.File.Copy(Fisier1, Fisier2, false);
                                                }
                                            }
                                            else
                                            {
                                                if (i > 1)
                                                {
                                                    if (is1 == true)
                                                    {
                                                        string Fisier1 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[0]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";
                                                        System.IO.File.Copy(Fisier1, Fisier2, false);
                                                        is1 = false;
                                                    }
                                                    else
                                                    {
                                                        string Fisier3 = Output_folder + _AGEN_mainform.dt_sheet_index.Rows[lista_generation[1]][_AGEN_mainform.Col_dwg_name].ToString() + ".dwg";
                                                        System.IO.File.Copy(Fisier3, Fisier2, false);
                                                        is1 = true;
                                                    }
                                                }
                                            }

                                            Display_dt.Rows.Add();
                                            Display_dt.Rows[Display_dt.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = Fisier2;
                                            Display_dt.Rows[Display_dt.Rows.Count - 1][_AGEN_mainform.Col_M1] = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_M1];
                                            Display_dt.Rows[Display_dt.Rows.Count - 1][_AGEN_mainform.Col_M2] = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_M2];
                                        }
                                    }




                                    Point3d PSpoint_prof = new Point3d(_AGEN_mainform.Vw_ps_prof_x, _AGEN_mainform.Vw_ps_prof_y, 0);

                                    System.Data.DataTable Data_table_poly = null;

                                    if (checkBox_profile.Checked == true)
                                    {
                                        Data_table_poly = create_profile_poly_definition(_AGEN_mainform.config_path);
                                    }

                                    if (checkBox_profile_band.Checked == true || checkBox_mult_vp_prof.Checked == true)
                                    {
                                        string fisier_prof_band = ProjFolder + _AGEN_mainform.band_prof_excel_name;
                                        if (System.IO.File.Exists(fisier_prof_band) == true)
                                        {
                                            _AGEN_mainform.Data_Table_profile_band = _AGEN_mainform.tpage_profdraw.Load_existing_profile_band_data(fisier_prof_band);
                                        }
                                    }



                                    for (int i = 0; i < lista_generation.Count; ++i)
                                    {
                                        List<Polyline> lista_poly = new List<Polyline>();

                                        string dwg_name = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_dwg_name]);
                                        string Fisier = Output_folder + dwg_name + ".dwg";
                                        using (Database Database2 = new Database(false, true))
                                        {

                                            Database2.ReadDwgFile(Fisier, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                            //System.IO.FileShare.ReadWrite, false, null);
                                            Database2.CloseInput(true);

                                            HostApplicationServices.WorkingDatabase = Database2;
                                            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_Main_Viewport, 4, false);


                                            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_North_Arrow, 7, true);

                                            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.layer_no_plot, 30, false);

                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                            {

                                                Functions.make_first_layout_active(Trans2, Database2);

                                                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                                                BtrecordPS.UpgradeOpen();
                                                Layout Layout1 = Functions.get_first_layout(Trans2, Database2);
                                                Layout1.UpgradeOpen();
                                                Layout1.LayoutName = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_dwg_name].ToString();



                                                #region VP multi profiles the same page
                                                if (checkBox_mult_vp_prof.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {

                                                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                                            {

                                                                _AGEN_mainform.Data_Table_regular_bands.Columns.Add("drafted", typeof(bool));
                                                                for (int k = 0; k < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++k)
                                                                {
                                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] = false;
                                                                }

                                                                for (int j = 0; j < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++j)
                                                                {
                                                                    string dwg_prof = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[j]["DwgNo"]);
                                                                    if (dwg_name.ToLower() == dwg_prof.ToLower())
                                                                    {

                                                                        double x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["x0"]);
                                                                        double y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["y0"]);
                                                                        double h1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["height"]);
                                                                        double l1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["length"]);

                                                                        ms_point = new Point3d(x0 + lr * l1 / 2, y0 + h1 / 2, 0);

                                                                        for (int k = 0; k < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++k)
                                                                        {
                                                                            if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["band_name"]) == _AGEN_mainform.tpage_viewport_settings.get_comboBox_bands_multiple_vp())
                                                                            {
                                                                                if ((bool)_AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] == false)
                                                                                {
                                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_ps_x"]);
                                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_ps_y"]);
                                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_width"]);
                                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_height"]);

                                                                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[k]["Custom_scale"] != DBNull.Value)
                                                                                    {
                                                                                        string str_scale = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["Custom_scale"]);
                                                                                        if (Functions.IsNumeric(str_scale) == true)
                                                                                        {
                                                                                            _AGEN_mainform.Vw_scale = Convert.ToDouble(str_scale);
                                                                                        }
                                                                                    }

                                                                                    Polyline rect1 = new Polyline();
                                                                                    rect1.AddVertexAt(0, new Point2d(ms_point.X - lr * (width1 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (height1 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                    rect1.AddVertexAt(1, new Point2d(ms_point.X - lr * (width1 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (height1 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                    rect1.AddVertexAt(2, new Point2d(ms_point.X + lr * (width1 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (height1 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                    rect1.AddVertexAt(3, new Point2d(ms_point.X + lr * (width1 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (height1 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                    rect1.Closed = true;

                                                                                    lista_poly.Add(rect1);

                                                                                    Point3d ps_point = new Point3d(xps, yps, 0);
                                                                                    Creaza_viewports_profile_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, width1, height1, _AGEN_mainform.Layer_name_profband_Viewport);
                                                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] = true;
                                                                                    k = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                _AGEN_mainform.Data_Table_regular_bands.Columns.Remove("drafted");
                                                            }


                                                        }
                                                    }
                                                }
                                                #endregion



                                                ms_point = new Point3d((double)_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_x], (double)_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_y], 0);
                                                double Twist = 2 * Math.PI - (double)_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_rot] * Math.PI / 180;

                                                if (_AGEN_mainform.Left_to_Right == false) Twist = Twist + Math.PI;

                                                double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_M1]);
                                                double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_AGEN_mainform.Col_M2]);


                                                #region plan view
                                                if (checkBox_plan_view.Checked == true)
                                                {
                                                    Viewport Viewport_main = Functions.Create_viewport(ms_point, ps_point_plan_view, _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_height, _AGEN_mainform.Vw_scale, Twist);
                                                    Viewport_main.Layer = _AGEN_mainform.Layer_name_Main_Viewport;
                                                    BtrecordPS.AppendEntity(Viewport_main);
                                                    Trans2.AddNewlyCreatedDBObject(Viewport_main, true);

                                                    ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                                                    oBJiD_COL.Add(Viewport_main.ObjectId);
                                                    DrawOrderTable DrawOrderTable2 = Trans2.GetObject(BtrecordPS.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                                                    DrawOrderTable2.MoveToBottom(oBJiD_COL);
                                                }
                                                #endregion


                                                #region extra viewports
                                                if (_AGEN_mainform.Data_Table_extra_mainVP != null && _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0)
                                                {
                                                    for (int j = 0; j < _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count; ++j)
                                                    {
                                                        if (_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"])) == true &&
                                                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"])) == true &&
                                                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"])) == true &&
                                                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"])) == true &&
                                                            _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"])) == true)
                                                        {

                                                            double scale1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"]);
                                                            double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"]);
                                                            double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"]);
                                                            double X_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"]);
                                                            double Y_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"]);

                                                            if (j == 0 && checkBoxx1.Checked == true)
                                                            {
                                                                Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_extra1_Viewport, 4, false);
                                                                Viewport Viewport_main1 = Functions.Create_viewport(ms_point, new Point3d(X_ps, Y_ps, 0), width1, height1, scale1, Twist);
                                                                Viewport_main1.Layer = _AGEN_mainform.Layer_name_extra1_Viewport;
                                                                BtrecordPS.AppendEntity(Viewport_main1);
                                                                Trans2.AddNewlyCreatedDBObject(Viewport_main1, true);
                                                            }

                                                            if (j == 1 && checkBoxx2.Checked == true)
                                                            {
                                                                Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_extra2_Viewport, 4, false);
                                                                Viewport Viewport_main2 = Functions.Create_viewport(ms_point, new Point3d(X_ps, Y_ps, 0), width1, height1, scale1, Twist);
                                                                Viewport_main2.Layer = _AGEN_mainform.Layer_name_extra2_Viewport;
                                                                BtrecordPS.AppendEntity(Viewport_main2);
                                                                Trans2.AddNewlyCreatedDBObject(Viewport_main2, true);
                                                            }

                                                            if (j == 2 && checkBoxx3.Checked == true)
                                                            {
                                                                Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_extra3_Viewport, 4, false);
                                                                Viewport Viewport_main3 = Functions.Create_viewport(ms_point, new Point3d(X_ps, Y_ps, 0), width1, height1, scale1, Twist);
                                                                Viewport_main3.Layer = _AGEN_mainform.Layer_name_extra3_Viewport;
                                                                BtrecordPS.AppendEntity(Viewport_main3);
                                                                Trans2.AddNewlyCreatedDBObject(Viewport_main3, true);
                                                            }


                                                            if (j == 3 && checkBoxx4.Checked == true)
                                                            {
                                                                Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_extra4_Viewport, 4, false);
                                                                Viewport Viewport_main4 = Functions.Create_viewport(ms_point, new Point3d(X_ps, Y_ps, 0), width1, height1, scale1, Twist);
                                                                Viewport_main4.Layer = _AGEN_mainform.Layer_name_extra4_Viewport;
                                                                BtrecordPS.AppendEntity(Viewport_main4);
                                                                Trans2.AddNewlyCreatedDBObject(Viewport_main4, true);
                                                            }


                                                            if (j == 4 && checkBoxx5.Checked == true)
                                                            {
                                                                Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_extra5_Viewport, 4, false);
                                                                Viewport Viewport_main5 = Functions.Create_viewport(ms_point, new Point3d(X_ps, Y_ps, 0), width1, height1, scale1, Twist);
                                                                Viewport_main5.Layer = _AGEN_mainform.Layer_name_extra5_Viewport;
                                                                BtrecordPS.AppendEntity(Viewport_main5);
                                                                Trans2.AddNewlyCreatedDBObject(Viewport_main5, true);
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion


                                                #region profile VP
                                                if (checkBox_profile.Checked == true)
                                                {
                                                    if (_AGEN_mainform.prof_width_lr > 0 && _AGEN_mainform.prof_texth > 0 && _AGEN_mainform.prof_x_left != _AGEN_mainform.prof_x_right &&
                                                        _AGEN_mainform.prof_x_right != -1.123 && _AGEN_mainform.prof_x_left != -1.123 && _AGEN_mainform.prof_y_down != -1.123 && _AGEN_mainform.prof_hexag != 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {

                                                            Creaza_viewports_profile(Fisier, Trans2, Database2, BtrecordPS, Data_table_poly, M1, M2, PSpoint_prof, lista_generation[i]);
                                                        }
                                                    }
                                                }
                                                #endregion


                                                double h_main = 2;

                                                #region VP profile_band
                                                if (checkBox_profile_band.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Vw_profband_height == 0)
                                                    {
                                                        MessageBox.Show("Profile band height = 0, verify your viewport settings!");
                                                        set_enable_true();
                                                        return;
                                                    }

                                                    if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {
                                                            for (int j = 0; j < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++j)
                                                            {
                                                                string dwg_prof = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[j]["DwgNo"]);
                                                                if (dwg_name.ToLower() == dwg_prof.ToLower())
                                                                {

                                                                    double x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["x0"]);
                                                                    double y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["y0"]);
                                                                    double h1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["height"]);
                                                                    double l1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["length"]);
                                                                    double staY = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["Sta_Y"]);
                                                                    double th = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["textH"]);


                                                                    h_main = _AGEN_mainform.Vw_profband_height;


                                                                    ms_point = new Point3d(x0 + lr * l1 / 2, y0 + h1 / 2, 0);

                                                                    Point3d ps_point = new Point3d(_AGEN_mainform.Vw_ps_profband_x, _AGEN_mainform.Vw_ps_profband_y, 0);


                                                                    Point3d ps_point_sta = new Point3d(_AGEN_mainform.Vw_ps_profband_x, _AGEN_mainform.Vw_ps_profband_y - h_main / 2, 0);


                                                                    Creaza_viewports_profile_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, _AGEN_mainform.Vw_width, h_main, _AGEN_mainform.Layer_name_profband_Viewport);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                double y_no_data_c = 0;

                                                #region no data viewport

                                                if (checkBox_no_data_band.Checked == true)
                                                {

                                                    if (h_nd == 0 || w_nd == 0)
                                                    {
                                                        MessageBox.Show("No data band height = 0, verify your viewport settings!");
                                                        set_enable_true();
                                                        return;
                                                    }

                                                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                    {
                                                        for (int j = 0; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                                        {
                                                            string dwg_si = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["DwgNo"]);
                                                            if (dwg_name.ToLower() == dwg_si.ToLower())
                                                            {

                                                                y_no_data_c = y_nd_ms - (j * sep_nd) / scale_nd - 0.5 * h_nd / scale_nd;


                                                                ms_point = new Point3d(x_nd_ms, y_no_data_c, 0);
                                                                Point3d ps_point = new Point3d(x_nd_ps, y_nd_ps, 0);
                                                                Creaza_viewports_no_data_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, w_nd, h_nd, scale_nd, _AGEN_mainform.Layer_name_no_data_band_Viewport);



                                                            }
                                                        }
                                                    }





                                                }


                                                #endregion


                                                #region vp ownership
                                                if (checkBox_ownership.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Vw_width > 0 && _AGEN_mainform.Vw_prop_height > 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {

                                                            Creaza_viewports_property_at_alignments_generation(Trans2, Database2, BtrecordPS,
                                                                 new Point3d(_AGEN_mainform.Point0_prop.X, (_AGEN_mainform.Point0_prop.Y - _AGEN_mainform.Vw_prop_height / 2) - lista_generation[i] * _AGEN_mainform.Band_Separation, 0), lista_generation[i]);
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region crossing VP
                                                if (checkBox_crossing.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Vw_width > 0 && _AGEN_mainform.Vw_cross_height > 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {

                                                            Creaza_viewports_crossing_at_alignments_generation(Trans2, Database2, BtrecordPS,
                                                                 new Point3d(_AGEN_mainform.Point0_cross.X, (_AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height / 2) - lista_generation[i] * _AGEN_mainform.Band_Separation, 0), lista_generation[i]);
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region materials vp
                                                if (checkBox_materials.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Vw_width > 0 && _AGEN_mainform.Vw_mat_height > 0)
                                                    {
                                                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                        {

                                                            Creaza_viewports_material_at_alignments_generation(Trans2, Database2, BtrecordPS,
                                                                  new Point3d(_AGEN_mainform.Point0_mat.X, (_AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height / 2) - lista_generation[i] * _AGEN_mainform.Band_Separation, 0), lista_generation[i]);


                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region north arrow
                                                if (_AGEN_mainform.NA_name != _AGEN_mainform.insertNAtoMS && _AGEN_mainform.NA_name != "")
                                                {
                                                    BlockTable BlockTable2 = Trans2.GetObject(Database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                                    if (BlockTable2 != null)
                                                    {
                                                        if (BlockTable2.Has(_AGEN_mainform.NA_name) == true)
                                                        {
                                                            BlockReference North_arrow = Functions.InsertBlock_with_multiple_atributes_with_database(Database2, BtrecordPS,
                                                                "", _AGEN_mainform.NA_name, new Point3d(_AGEN_mainform.NA_x, _AGEN_mainform.NA_y, 0),
                                                                _AGEN_mainform.NA_scale, Twist, _AGEN_mainform.Layer_North_Arrow, new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region custom bands
                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                                                {
                                                    Point3d ms_pt = new Point3d();
                                                    Point3d ps_pt = new Point3d();

                                                    double twist1 = 0;

                                                    if (checkBox1.Visible == true)
                                                    {
                                                        if (checkBox1.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 1)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_y"]);



                                                                    double y1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_y"]) - 0.5 * height1 / scale_cust - lista_generation[i] * _AGEN_mainform.Band_Separation / scale_cust;

                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_x"]), y1, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, scale_cust, twist1);



                                                                }


                                                            }
                                                        }
                                                    }

                                                    if (checkBox2.Visible == true)
                                                    {
                                                        if (checkBox2.Checked == true)
                                                        {

                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 2)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[1]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (checkBox3.Visible == true)
                                                    {
                                                        if (checkBox3.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 3)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[2]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (checkBox4.Visible == true)
                                                    {
                                                        if (checkBox4.Checked == true)
                                                        {

                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 4)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[3]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }

                                                        }
                                                    }

                                                    if (checkBox5.Visible == true)
                                                    {

                                                        if (checkBox5.Checked == true)
                                                        {

                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 5)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[4]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (checkBox6.Visible == true)
                                                    {
                                                        if (checkBox6.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 6)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[5]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }

                                                        }
                                                    }


                                                    if (checkBox7.Visible == true)
                                                    {
                                                        if (checkBox7.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 7)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[6]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox8.Visible == true)
                                                    {
                                                        if (checkBox8.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 8)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[7]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox9.Visible == true)
                                                    {
                                                        if (checkBox9.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 9)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[8]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox10.Visible == true)
                                                    {
                                                        if (checkBox10.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 10)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[9]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox11.Visible == true)
                                                    {
                                                        if (checkBox11.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[10]["band_name"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_height"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_width"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_x"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_y"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_x"] != DBNull.Value &&
                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_y"] != DBNull.Value)
                                                            {
                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["band_name"]) + "VP";

                                                                double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_height"]);
                                                                double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_width"]);

                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_x"]);
                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_y"]);


                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_x"]),
                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                    lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                            }
                                                        }
                                                    }


                                                    if (checkBox12.Visible == true)
                                                    {
                                                        if (checkBox12.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 12)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[11]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox13.Visible == true)
                                                    {
                                                        if (checkBox13.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 13)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[12]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox14.Visible == true)
                                                    {
                                                        if (checkBox14.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 14)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[13]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox15.Visible == true)
                                                    {
                                                        if (checkBox15.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 15)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[14]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }


                                                    if (checkBox16.Visible == true)
                                                    {
                                                        if (checkBox16.Checked == true)
                                                        {
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 16)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[15]["band_name"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_height"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_width"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_y"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_x"] != DBNull.Value &&
                                                                    _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_y"] != DBNull.Value)
                                                                {
                                                                    string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["band_name"]) + "VP";

                                                                    double height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_height"]);
                                                                    double width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_width"]);

                                                                    double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_x"]);
                                                                    double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_y"]);


                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_x"]),
                                                                                        Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_y"]) - 0.5 * height1 -
                                                                                        lista_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                                                    ps_pt = new Point3d(xps, yps, 0);

                                                                    Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, twist1);

                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion


                                                #region matchline paperspace
                                                if (_AGEN_mainform.Matchline_BlockName_in_PaperSpace != "")
                                                {
                                                    BlockTable BlockTable2 = Trans2.GetObject(Database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                                    if (BlockTable2 != null)
                                                    {
                                                        if (BlockTable2.Has(_AGEN_mainform.Matchline_BlockName_in_PaperSpace) == true)
                                                        {
                                                            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_Matchline_PaperSpace, 1, true);

                                                            double dispm1 = Functions.Station_equation_ofV2(M1, _AGEN_mainform.dt_station_equation);
                                                            double dispm2 = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);

                                                            string StM1 = Functions.Get_chainage_from_double(dispm1, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                            string StM2 = Functions.Get_chainage_from_double(dispm2, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                            string mile1 = Functions.Get_String_Rounded(dispm1 / 5280, 1);
                                                            string mile2 = Functions.Get_String_Rounded(dispm2 / 5280, 1);

                                                            string Prev_file = "BEGIN STA.";
                                                            string Next_file = "END STA.";

                                                            double ml_width = 0;

                                                            if (_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]]["Width"] != DBNull.Value &&
                                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]]["Width"])) == true)
                                                            {
                                                                ml_width = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i]]["Width"]) * _AGEN_mainform.Vw_scale;
                                                            }

                                                            if (_AGEN_mainform.dt_sheet_index.Rows.Count > 1)
                                                            {
                                                                if (lista_generation[i] < _AGEN_mainform.dt_sheet_index.Rows.Count - 1)
                                                                {
                                                                    if (_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i] + 1][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                                                    {
                                                                        Next_file = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i] + 1][_AGEN_mainform.Col_dwg_name].ToString();
                                                                    }
                                                                }

                                                                if (lista_generation[i] > 0)
                                                                {
                                                                    if (_AGEN_mainform.dt_sheet_index.Rows[lista_generation[i] - 1][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                                                    {
                                                                        Prev_file = _AGEN_mainform.dt_sheet_index.Rows[lista_generation[i] - 1][_AGEN_mainform.Col_dwg_name].ToString();
                                                                    }
                                                                }
                                                            }

                                                            double y = _AGEN_mainform.Vw_ps_y - _AGEN_mainform.Vw_height / 2;
                                                            double x1 = _AGEN_mainform.Vw_ps_x - ml_width / 2;
                                                            double x2 = _AGEN_mainform.Vw_ps_x + ml_width / 2;
                                                            if (_AGEN_mainform.Left_to_Right == false)
                                                            {
                                                                double t = x1;
                                                                x1 = x2;
                                                                x2 = t;
                                                            }


                                                            System.Collections.Specialized.StringCollection col_atr_left = new System.Collections.Specialized.StringCollection();
                                                            System.Collections.Specialized.StringCollection col_val_left = new System.Collections.Specialized.StringCollection();

                                                            col_atr_left.Add("ATR_1");
                                                            col_val_left.Add(mile1);
                                                            col_atr_left.Add("ATR_2");
                                                            col_val_left.Add(Prev_file);
                                                            col_atr_left.Add("ATR_3");
                                                            col_val_left.Add(StM1);

                                                            col_atr_left.Add("ATR_4");
                                                            col_val_left.Add("");
                                                            col_atr_left.Add("ATR_5");
                                                            col_val_left.Add("");

                                                            BlockReference ML_PS_left = Functions.InsertBlock_with_multiple_atributes_with_database(Database2, BtrecordPS,
                                                                "", _AGEN_mainform.Matchline_BlockName_in_PaperSpace, new Point3d(x1, y, 0),
                                                                1, 0, _AGEN_mainform.Layer_Matchline_PaperSpace, col_atr_left, col_val_left);
                                                            Functions.Stretch_block(ML_PS_left, "Distance1", _AGEN_mainform.Vw_height);

                                                            System.Collections.Specialized.StringCollection col_atr_right = new System.Collections.Specialized.StringCollection();
                                                            System.Collections.Specialized.StringCollection col_val_right = new System.Collections.Specialized.StringCollection();

                                                            col_atr_right.Add("ATR_1");
                                                            col_val_right.Add(mile2);
                                                            col_atr_right.Add("ATR_5");
                                                            col_val_right.Add(Next_file);
                                                            col_atr_right.Add("ATR_4");
                                                            col_val_right.Add(StM2);

                                                            col_atr_right.Add("ATR_2");
                                                            col_val_right.Add("");
                                                            col_atr_right.Add("ATR_3");
                                                            col_val_right.Add("");

                                                            BlockReference ML_PS_right = Functions.InsertBlock_with_multiple_atributes_with_database(Database2, BtrecordPS,
                                                                "", _AGEN_mainform.Matchline_BlockName_in_PaperSpace, new Point3d(x2, y, 0),
                                                                1, 0, _AGEN_mainform.Layer_Matchline_PaperSpace, col_atr_right, col_val_right);
                                                            Functions.Stretch_block(ML_PS_right, "Distance1", _AGEN_mainform.Vw_height);

                                                        }
                                                    }
                                                }
                                                #endregion


                                                Trans2.Commit();

                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                Database2.SaveAs(Fisier, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);

                                                #region profile band vp rectangles
                                                if (checkBox_profile_band.Checked == true)
                                                {
                                                    Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                                    Polyline polyVP1 = new Polyline();

                                                    polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                    polyVP1.ColorIndex = 1;

                                                    polyVP1.AddVertexAt(0, new Point2d(ms_point.X - lr * (_AGEN_mainform.Vw_width / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                    polyVP1.AddVertexAt(1, new Point2d(ms_point.X - lr * (_AGEN_mainform.Vw_width / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                    polyVP1.AddVertexAt(2, new Point2d(ms_point.X + lr * (_AGEN_mainform.Vw_width / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                    polyVP1.AddVertexAt(3, new Point2d(ms_point.X + lr * (_AGEN_mainform.Vw_width / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                    polyVP1.Closed = true;

                                                    Btrecord.AppendEntity(polyVP1);
                                                    Trans1.AddNewlyCreatedDBObject(polyVP1, true);

                                                }
                                                #endregion

                                                #region MULTIPLE profile vp rectangles
                                                if (checkBox_mult_vp_prof.Checked == true)
                                                {
                                                    if (lista_poly.Count > 0)
                                                    {
                                                        Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);

                                                        for (int m = 0; m < lista_poly.Count; ++m)
                                                        {

                                                            Polyline polyVP1 = new Polyline();
                                                            polyVP1 = lista_poly[m];
                                                            polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                            polyVP1.ColorIndex = 1;
                                                            Btrecord.AppendEntity(polyVP1);
                                                            Trans1.AddNewlyCreatedDBObject(polyVP1, true);
                                                        }
                                                    }
                                                }
                                                #endregion



                                                #region nodata band vp rectangles
                                                if (checkBox_no_data_band.Checked == true)
                                                {
                                                    Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                                    Polyline polyVP1 = new Polyline();

                                                    polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                    polyVP1.ColorIndex = 1;

                                                    polyVP1.AddVertexAt(0, new Point2d(x_nd_ms - (w_nd / 2) / scale_nd, y_no_data_c + 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                    polyVP1.AddVertexAt(1, new Point2d(x_nd_ms + (w_nd / 2) / scale_nd, y_no_data_c + 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                    polyVP1.AddVertexAt(2, new Point2d(x_nd_ms + (w_nd / 2) / scale_nd, y_no_data_c - 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                    polyVP1.AddVertexAt(3, new Point2d(x_nd_ms - (w_nd / 2) / scale_nd, y_no_data_c - 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                    polyVP1.Closed = true;

                                                    Btrecord.AppendEntity(polyVP1);
                                                    Trans1.AddNewlyCreatedDBObject(polyVP1, true);

                                                }
                                                #endregion

                                            }
                                        }
                                    }



                                    dataGridView_align_created.DataSource = Display_dt;
                                    dataGridView_align_created.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                                    dataGridView_align_created.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                    dataGridView_align_created.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                                    dataGridView_align_created.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                                    dataGridView_align_created.DefaultCellStyle.ForeColor = Color.White;
                                    dataGridView_align_created.EnableHeadersVisualStyles = false;
                                }
                                #endregion


                                #region update file with new vp

                                if (Creaza_new_file == false && (
                                        checkBox_ownership.Checked == true
                                        || checkBox_plan_view.Checked == true
                                        || checkBox_materials.Checked == true
                                        || checkBox_crossing.Checked == true
                                        || checkBox_profile.Checked == true
                                        || checkBox1.Checked == true
                                        || checkBox2.Checked == true
                                        || checkBox3.Checked == true
                                        || checkBox4.Checked == true
                                        || checkBox5.Checked == true
                                        || checkBox6.Checked == true
                                        || checkBox7.Checked == true
                                        || checkBox8.Checked == true
                                        || checkBox9.Checked == true
                                        || checkBox10.Checked == true
                                        || checkBox11.Checked == true
                                        || checkBox12.Checked == true
                                        || checkBox13.Checked == true
                                        || checkBox14.Checked == true
                                        || checkBox15.Checked == true
                                        || checkBox16.Checked == true
                                        || checkBox_slope_band.Checked == true
                                        || checkBox_profile_band.Checked == true
                                        || checkBox_xref_clip.Checked == true
                                        || checkBox_tblk.Checked == true
                                        || checkBox_no_data_band.Checked == true
                                        || checkBoxx1.Checked == true
                                        || checkBoxx2.Checked == true
                                        || checkBoxx3.Checked == true
                                        || checkBoxx4.Checked == true
                                        || checkBoxx5.Checked == true
                                        || checkBox_mult_vp_prof.Checked == true))

                                {
                                    System.Data.DataTable Data_table_poly = null;


                                    if (checkBox_profile.Checked == true)
                                    {
                                        Data_table_poly = create_profile_poly_definition(_AGEN_mainform.config_path);
                                    }

                                    List<string> Lista_bl = new List<string>();
                                    Lista_bl.Add("spire_matchline_left");
                                    Lista_bl.Add("spire_matchline_left1");
                                    Lista_bl.Add("spire_matchline_right");
                                    Lista_bl.Add("spire_matchline_right1");
                                    Lista_bl.Add("north_arrow");

                                    if (checkBox_profile_band.Checked == true)
                                    {
                                        _AGEN_mainform.tpage_profdraw.button_load_data_for_profile_band_Click(sender, e);
                                    }

                                    for (int i = 0; i < Display_dt.Rows.Count; ++i)
                                    {
                                        List<Polyline> lista_poly = new List<Polyline>();
                                        if (Display_dt.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                        {
                                            string file1 = Display_dt.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                            string nume_fara_ext = System.IO.Path.GetFileNameWithoutExtension(file1);

                                            if (System.IO.File.Exists(file1) == true)
                                            {
                                                Point3d ms_pt = new Point3d();
                                                Point3d ps_pt = new Point3d();
                                                double width1 = 0;
                                                double height1 = 0;
                                                int si_index = -1;
                                                int prof_band_index = -1;


                                                if (checkBox_profile_band.Checked == true)
                                                {
                                                    if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                                                    {
                                                        for (int j = 0; j < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++j)
                                                        {
                                                            string si_name = _AGEN_mainform.Data_Table_profile_band.Rows[j]["DwgNo"].ToString();
                                                            if (si_name.ToLower() == nume_fara_ext.ToLower())
                                                            {
                                                                prof_band_index = j;
                                                                j = _AGEN_mainform.Data_Table_profile_band.Rows.Count;

                                                            }
                                                        }
                                                    }
                                                    if (prof_band_index == -1)
                                                    {
                                                        MessageBox.Show("Sheet index data does not match profile band data.\r\nOperation aborted.", "agen", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                        _AGEN_mainform.tpage_processing.Hide();
                                                        Freeze_operations = false;
                                                        return;
                                                    }
                                                }


                                                double y_no_data_c = 0;

                                                for (int j = 0; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                                {
                                                    string si_name = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name].ToString();
                                                    if (si_name.ToLower() == nume_fara_ext.ToLower())
                                                    {
                                                        si_index = j;
                                                        y_no_data_c = y_nd_ms - (j * sep_nd) / scale_nd - 0.5 * h_nd / scale_nd;
                                                        j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                                    }
                                                }

                                                if (checkBox_mult_vp_prof.Checked == true)
                                                {
                                                    string fisier_prof_band = ProjFolder + _AGEN_mainform.band_prof_excel_name;
                                                    if (System.IO.File.Exists(fisier_prof_band) == true)
                                                    {
                                                        _AGEN_mainform.Data_Table_profile_band = _AGEN_mainform.tpage_profdraw.Load_existing_profile_band_data(fisier_prof_band);
                                                    }
                                                }

                                                double h_main = 10;

                                                double l_main = _AGEN_mainform.Vw_width;


                                                if (si_index >= 0 || prof_band_index >= 0 || checkBox_mult_vp_prof.Checked == true)
                                                {
                                                    using (Database Database2 = new Database(false, true))
                                                    {
                                                        Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                                        //System.IO.FileShare.ReadWrite, false, null);
                                                        Database2.CloseInput(true);

                                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                                        {
                                                            Functions.make_first_layout_active(Trans2, Database2);
                                                            BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                                                            BtrecordPS.UpgradeOpen();
                                                            Layout Layout1 = Functions.get_first_layout(Trans2, Database2);
                                                            Layout1.UpgradeOpen();
                                                            //Layout1.LayoutName = _AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_dwg_name].ToString();

                                                            #region delete viewport
                                                            if (checkBox_delete_vp.Checked == true)
                                                            {

                                                                Delete_viewport_on_existing_alignment(Database2);

                                                            }
                                                            #endregion


                                                            #region VP multi profiles the same page
                                                            if (checkBox_mult_vp_prof.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                                                                {
                                                                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                                    {

                                                                        if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                                                        {

                                                                            if (_AGEN_mainform.Data_Table_regular_bands.Columns.Contains("drafted") == false) _AGEN_mainform.Data_Table_regular_bands.Columns.Add("drafted", typeof(bool));
                                                                            for (int k = 0; k < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++k)
                                                                            {
                                                                                _AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] = false;
                                                                            }

                                                                            foreach (ObjectId odid in BtrecordPS)
                                                                            {
                                                                                Entity ent1 = Trans2.GetObject(odid, OpenMode.ForRead) as Entity;

                                                                                if (ent1 != null)
                                                                                {
                                                                                    if (ent1.Layer.ToLower() == _AGEN_mainform.Layer_name_profband_Viewport.ToLower())
                                                                                    {
                                                                                        ent1.UpgradeOpen();
                                                                                        ent1.Erase();
                                                                                    }
                                                                                }

                                                                            }


                                                                            for (int j = 0; j < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++j)
                                                                            {
                                                                                string dwg_prof = Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[j]["DwgNo"]);
                                                                                if (nume_fara_ext.ToLower() == dwg_prof.ToLower())
                                                                                {

                                                                                    double x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["x0"]);
                                                                                    double y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["y0"]);
                                                                                    double h1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["height"]);
                                                                                    double l1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[j]["length"]);

                                                                                    ms_point = new Point3d(x0 + lr * l1 / 2, y0 + h1 / 2, 0);

                                                                                    for (int k = 0; k < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++k)
                                                                                    {
                                                                                        if (Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["band_name"]) == _AGEN_mainform.tpage_viewport_settings.get_comboBox_bands_multiple_vp())
                                                                                        {
                                                                                            if ((bool)_AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] == false)
                                                                                            {
                                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_ps_x"]);
                                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_ps_y"]);
                                                                                                double width2 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_width"]);
                                                                                                double height2 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["viewport_height"]);

                                                                                                if (_AGEN_mainform.Data_Table_regular_bands.Rows[k]["Custom_scale"] != DBNull.Value)
                                                                                                {
                                                                                                    string str_scale = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[k]["Custom_scale"]);
                                                                                                    if (Functions.IsNumeric(str_scale) == true)
                                                                                                    {
                                                                                                        _AGEN_mainform.Vw_scale = Convert.ToDouble(str_scale);
                                                                                                    }
                                                                                                }

                                                                                                Polyline rect1 = new Polyline();
                                                                                                rect1.AddVertexAt(0, new Point2d(ms_point.X - lr * (width2 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (height2 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                                rect1.AddVertexAt(1, new Point2d(ms_point.X - lr * (width2 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (height2 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                                rect1.AddVertexAt(2, new Point2d(ms_point.X + lr * (width2 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (height2 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                                rect1.AddVertexAt(3, new Point2d(ms_point.X + lr * (width2 / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (height2 / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                                                                rect1.Closed = true;

                                                                                                lista_poly.Add(rect1);

                                                                                                Point3d ps_point = new Point3d(xps, yps, 0);
                                                                                                Creaza_viewports_profile_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, width2, height2, _AGEN_mainform.Layer_name_profband_Viewport);

                                                                                                _AGEN_mainform.Data_Table_regular_bands.Rows[k]["drafted"] = true;
                                                                                                k = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }

                                                                            _AGEN_mainform.Data_Table_regular_bands.Columns.Remove("drafted");
                                                                        }


                                                                    }
                                                                }
                                                            }
                                                            #endregion



                                                            #region VP plan view
                                                            if (checkBox_plan_view.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {

                                                                ms_pt = new Point3d((double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_x], (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_y], 0);
                                                                double twist1 = 2 * Math.PI - (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_rot] * Math.PI / 180;
                                                                if (_AGEN_mainform.Left_to_Right == false) twist1 = twist1 + Math.PI;
                                                                Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_Main_Viewport, ms_pt,
                                                                                                                                new Point3d(_AGEN_mainform.Vw_ps_x, _AGEN_mainform.Vw_ps_y, 0),
                                                                                                                                _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_height, _AGEN_mainform.Vw_scale, twist1);
                                                            }
                                                            #endregion

                                                            #region extra viewports
                                                            if (_AGEN_mainform.Data_Table_extra_mainVP != null && _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count > 0 && checkBox_delete_vp.Checked == false)
                                                            {

                                                                ms_pt = new Point3d((double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_x], (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_y], 0);
                                                                double twist1 = 2 * Math.PI - (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_rot] * Math.PI / 180;
                                                                if (_AGEN_mainform.Left_to_Right == false) twist1 = twist1 + Math.PI;

                                                                for (int j = 0; j < _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count; ++j)
                                                                {
                                                                    if (_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"])) == true &&
                                                                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"])) == true &&
                                                                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"])) == true &&
                                                                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"])) == true &&
                                                                        _AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"])) == true)
                                                                    {

                                                                        double scale1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["Custom_scale"]);
                                                                        double widthx1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_width"]);
                                                                        double heightx1 = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_height"]);
                                                                        double X_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_x"]);
                                                                        double Y_ps = Convert.ToDouble(_AGEN_mainform.Data_Table_extra_mainVP.Rows[j]["viewport_ps_y"]);

                                                                        if (j == 0 && checkBoxx1.Checked == true)
                                                                        {
                                                                            Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_extra1_Viewport, ms_pt, new Point3d(X_ps, Y_ps, 0), widthx1, heightx1, scale1, twist1);
                                                                        }
                                                                        if (j == 1 && checkBoxx2.Checked == true)
                                                                        {
                                                                            Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_extra2_Viewport, ms_pt, new Point3d(X_ps, Y_ps, 0), widthx1, heightx1, scale1, twist1);
                                                                        }
                                                                        if (j == 2 && checkBoxx3.Checked == true)
                                                                        {
                                                                            Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_extra3_Viewport, ms_pt, new Point3d(X_ps, Y_ps, 0), widthx1, heightx1, scale1, twist1);
                                                                        }
                                                                        if (j == 3 && checkBoxx4.Checked == true)
                                                                        {
                                                                            Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_extra4_Viewport, ms_pt, new Point3d(X_ps, Y_ps, 0), widthx1, heightx1, scale1, twist1);
                                                                        }
                                                                        if (j == 4 && checkBoxx5.Checked == true)
                                                                        {
                                                                            Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_extra4_Viewport, ms_pt, new Point3d(X_ps, Y_ps, 0), widthx1, heightx1, scale1, twist1);
                                                                        }

                                                                    }
                                                                }
                                                            }
                                                            #endregion

                                                            #region XREF clip

                                                            if (checkBox_xref_clip.Checked == true)
                                                            {

                                                                BlockTableRecord BtrecordMS = Functions.get_modelspace(Trans2, Database2);

                                                                BlockTable BlockTable2 = Trans2.GetObject(Database2.BlockTableId, OpenMode.ForRead) as BlockTable;
                                                                if (BlockTable2.Has(textBox_xref_name.Text) == true)
                                                                {

                                                                    BlockTableRecord xref_rec = Trans2.GetObject(BlockTable2[textBox_xref_name.Text], OpenMode.ForRead) as BlockTableRecord;
                                                                    if (xref_rec.IsFromExternalReference == true)
                                                                    {
                                                                        foreach (ObjectId id1 in BtrecordMS)
                                                                        {
                                                                            BlockReference xref1 = Trans2.GetObject(id1, OpenMode.ForRead) as BlockReference;
                                                                            if (xref1 != null)
                                                                            {
                                                                                string numeXref = xref1.Name;
                                                                                if (numeXref == textBox_xref_name.Text)
                                                                                {
                                                                                    ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_x]),
                                                                                                        Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_y]), 0);

                                                                                    double Rotation = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_rot]) * Math.PI / 180;

                                                                                    Polyline poly_clip = _AGEN_mainform.tpage_sheetindex.creaza_rectangle_from_one_point(ms_pt, Rotation, _AGEN_mainform.Vw_width / _AGEN_mainform.Vw_scale,
                                                                                                                                                                                _AGEN_mainform.Vw_height / _AGEN_mainform.Vw_scale, 1);

                                                                                    // Set the clipping boundary and enable it
                                                                                    using (Autodesk.AutoCAD.DatabaseServices.Filters.SpatialFilter filter = new Autodesk.AutoCAD.DatabaseServices.Filters.SpatialFilter())
                                                                                    {

                                                                                        Point2dCollection ptCol = new Point2dCollection();

                                                                                        for (int n = 0; n < poly_clip.NumberOfVertices; ++n)
                                                                                        {
                                                                                            ptCol.Add(poly_clip.GetPoint2dAt(n));
                                                                                        }

                                                                                        // Define the normal and elevation for the clipping boundary 
                                                                                        Vector3d Zaxis;
                                                                                        double Zclip = 0;

                                                                                        if (Database2.TileMode == true)
                                                                                        {
                                                                                            Zaxis = Database2.Ucsxdir.CrossProduct(Database2.Ucsydir);
                                                                                            Zclip = Database2.Elevation;
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            Zaxis = Database2.Pucsxdir.CrossProduct(Database2.Pucsydir);
                                                                                            Zclip = Database2.Pelevation;
                                                                                        }

                                                                                        Zclip = 30000;


                                                                                        Autodesk.AutoCAD.DatabaseServices.Filters.SpatialFilterDefinition filterDef =
                                                                                            new Autodesk.AutoCAD.DatabaseServices.Filters.SpatialFilterDefinition(ptCol, Zaxis, Zclip, 0, 0, true);
                                                                                        filter.Definition = filterDef;

                                                                                        // Define the name of the extension dictionary and entry name
                                                                                        string dictName = "ACAD_FILTER";
                                                                                        string spName = "SPATIAL";

                                                                                        // Check to see if the Extension Dictionary exists, if not create it
                                                                                        if (xref1.ExtensionDictionary.IsNull)
                                                                                        {
                                                                                            xref1.UpgradeOpen();
                                                                                            xref1.CreateExtensionDictionary();
                                                                                            xref1.DowngradeOpen();
                                                                                        }

                                                                                        // Open the Extension Dictionary for write
                                                                                        DBDictionary extDict = Trans2.GetObject(xref1.ExtensionDictionary, OpenMode.ForWrite) as DBDictionary;

                                                                                        // Check to see if the dictionary for clipped boundaries exists, 
                                                                                        // and add the spatial filter to the dictionary
                                                                                        if (extDict.Contains(dictName))
                                                                                        {
                                                                                            DBDictionary filterDict = Trans2.GetObject(extDict.GetAt(dictName), OpenMode.ForWrite) as DBDictionary;

                                                                                            if (filterDict.Contains(spName))
                                                                                            {
                                                                                                filterDict.Remove(spName);
                                                                                            }

                                                                                            filterDict.SetAt(spName, filter);
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            using (DBDictionary filterDict = new DBDictionary())
                                                                                            {
                                                                                                extDict.SetAt(dictName, filterDict);

                                                                                                Trans2.AddNewlyCreatedDBObject(filterDict, true);
                                                                                                filterDict.SetAt(spName, filter);
                                                                                            }
                                                                                        }

                                                                                        // Append the spatial filter to the drawing
                                                                                        Trans2.AddNewlyCreatedDBObject(filter, true);
                                                                                    }



                                                                                }
                                                                            }
                                                                        }
                                                                    }



                                                                }






                                                            }

                                                            #endregion

                                                            #region VP ownership
                                                            if (checkBox_ownership.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                ms_pt = new Point3d(_AGEN_mainform.Point0_prop.X, (_AGEN_mainform.Point0_prop.Y - _AGEN_mainform.Vw_prop_height / 2) - si_index * _AGEN_mainform.Band_Separation, 0);
                                                                Point3d center_ps = new Point3d(_AGEN_mainform.Vw_ps_prop_x, _AGEN_mainform.Vw_ps_prop_y, 0);
                                                                Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_ownership_Viewport, ms_pt, center_ps, _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_prop_height, 1, 0);
                                                            }
                                                            #endregion

                                                            #region VP materials
                                                            if (checkBox_materials.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                ms_pt = new Point3d(_AGEN_mainform.Point0_mat.X, (_AGEN_mainform.Point0_mat.Y - _AGEN_mainform.Vw_mat_height / 2) - si_index * _AGEN_mainform.Band_Separation, 0);
                                                                Point3d center_ps = new Point3d(_AGEN_mainform.Vw_ps_mat_x, _AGEN_mainform.Vw_ps_mat_y, 0);
                                                                Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_material_Viewport, ms_pt, center_ps, _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_mat_height, 1, 0);

                                                            }
                                                            #endregion


                                                            #region VP crossing
                                                            if (checkBox_crossing.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                ms_pt = new Point3d(_AGEN_mainform.Point0_cross.X, (_AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height / 2) - si_index * _AGEN_mainform.Band_Separation, 0);
                                                                Point3d center_ps = new Point3d(_AGEN_mainform.Vw_ps_cross_x, _AGEN_mainform.Vw_ps_cross_y, 0);
                                                                Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_crossing_Viewport, ms_pt, center_ps, _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_cross_height, 1, 0);
                                                            }
                                                            #endregion
                                                            #region VP profile
                                                            if (checkBox_profile.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                if (_AGEN_mainform.prof_width_lr > 0 && _AGEN_mainform.prof_texth > 0 && _AGEN_mainform.prof_x_left != _AGEN_mainform.prof_x_right &&
                                                                    _AGEN_mainform.prof_x_right != -1.123 && _AGEN_mainform.prof_x_left != -1.123 && _AGEN_mainform.prof_y_down != -1.123 && _AGEN_mainform.prof_hexag != 0)
                                                                {

                                                                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                                    {

                                                                        ms_point = new Point3d((double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_x], (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_y], 0);
                                                                        double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_M1]);
                                                                        double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_M2]);
                                                                        Point3d PSpoint_prof = new Point3d(_AGEN_mainform.Vw_ps_prof_x, _AGEN_mainform.Vw_ps_prof_y, 0);
                                                                        string Layer_name_prof_main_viewport_old = "VP_Prof_ON";
                                                                        string Layer_name_prof_viewport_old = "VP_Prof_OFF";

                                                                        foreach (ObjectId odid in BtrecordPS)
                                                                        {
                                                                            Entity ent1 = Trans2.GetObject(odid, OpenMode.ForRead) as Entity;
                                                                            if (checkBox_profile.Checked == true)
                                                                            {
                                                                                if (ent1 != null)
                                                                                {
                                                                                    if (ent1.Layer.ToLower() == _AGEN_mainform.Layer_name_prof_side_viewport.ToLower() || ent1.Layer.ToLower() == _AGEN_mainform.Layer_name_prof_main_viewport.ToLower()
                                                                                        || ent1.Layer.ToLower() == Layer_name_prof_viewport_old.ToLower() || ent1.Layer.ToLower() == Layer_name_prof_main_viewport_old.ToLower())
                                                                                    {
                                                                                        ent1.UpgradeOpen();
                                                                                        ent1.Erase();
                                                                                    }
                                                                                }
                                                                            }


                                                                        }
                                                                        Creaza_viewports_profile(file1, Trans2, Database2, BtrecordPS, Data_table_poly, M1, M2, PSpoint_prof, si_index);

                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                            #region VP profile_band
                                                            if (checkBox_profile_band.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {



                                                                if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                                                                {

                                                                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                                    {
                                                                        double x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["x0"]);
                                                                        double y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["y0"]);
                                                                        double h1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["height"]);
                                                                        double l1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["length"]);
                                                                        double staY = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["Sta_Y"]);
                                                                        double th = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[prof_band_index]["textH"]);


                                                                        h_main = _AGEN_mainform.Vw_profband_height;


                                                                        ms_point = new Point3d(x0 + lr * l1 / 2, y0 + h1 / 2, 0);

                                                                        Point3d ps_point = new Point3d(_AGEN_mainform.Vw_ps_profband_x, _AGEN_mainform.Vw_ps_profband_y, 0);


                                                                        Point3d ps_point_sta = new Point3d(_AGEN_mainform.Vw_ps_profband_x, _AGEN_mainform.Vw_ps_profband_y - h_main / 2, 0);



                                                                        foreach (ObjectId odid in BtrecordPS)
                                                                        {
                                                                            Entity ent1 = Trans2.GetObject(odid, OpenMode.ForRead) as Entity;

                                                                            if (ent1 != null)
                                                                            {
                                                                                if (ent1.Layer.ToLower() == _AGEN_mainform.Layer_name_profband_Viewport.ToLower())
                                                                                {
                                                                                    ent1.UpgradeOpen();
                                                                                    ent1.Erase();
                                                                                }
                                                                            }

                                                                        }
                                                                        Creaza_viewports_profile_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, l_main, h_main, _AGEN_mainform.Layer_name_profband_Viewport);
                                                                    }
                                                                }
                                                            }
                                                            #endregion

                                                            #region VP no data band
                                                            if (checkBox_no_data_band.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {




                                                                if (h_nd == 0 || w_nd == 0)
                                                                {
                                                                    MessageBox.Show("No data band height = 0, verify your viewport settings!");
                                                                    set_enable_true();
                                                                    return;
                                                                }


                                                                foreach (ObjectId odid in BtrecordPS)
                                                                {
                                                                    Entity ent1 = Trans2.GetObject(odid, OpenMode.ForRead) as Entity;

                                                                    if (ent1 != null)
                                                                    {
                                                                        if (ent1.Layer.ToLower() == _AGEN_mainform.Layer_name_no_data_band_Viewport.ToLower())
                                                                        {
                                                                            ent1.UpgradeOpen();
                                                                            ent1.Erase();
                                                                        }
                                                                    }

                                                                }




                                                                ms_point = new Point3d(x_nd_ms, y_no_data_c, 0);
                                                                Point3d ps_point = new Point3d(x_nd_ps, y_nd_ps, 0);
                                                                Creaza_viewports_no_data_band(Trans2, Database2, BtrecordPS, ms_point, ps_point, w_nd, h_nd, scale_nd, _AGEN_mainform.Layer_name_no_data_band_Viewport);







                                                            }
                                                            #endregion

                                                            #region VP TBLK
                                                            if (checkBox_tblk.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                ms_pt = new Point3d(_AGEN_mainform.Point0_tblk.X, (_AGEN_mainform.Point0_tblk.Y - _AGEN_mainform.Vw_tblk_height / 2) - si_index * _AGEN_mainform.tblk_separation, 0);

                                                                if (_AGEN_mainform.tblk_twist == 90 * Math.PI / 180)
                                                                {
                                                                    ms_pt = new Point3d(_AGEN_mainform.Point0_tblk.X, (_AGEN_mainform.Point0_tblk.Y - _AGEN_mainform.Vw_tblk_width / 2) - si_index * _AGEN_mainform.tblk_separation, 0);
                                                                }

                                                                Point3d center_ps = new Point3d(_AGEN_mainform.Vw_ps_tblk_x, _AGEN_mainform.Vw_ps_tblk_y, 0);
                                                                Creaza_viewport_on_alignment_on_existing(Database2, _AGEN_mainform.Layer_name_tblk_Viewport, ms_pt, center_ps, _AGEN_mainform.Vw_tblk_width, _AGEN_mainform.Vw_tblk_height, 1, _AGEN_mainform.tblk_twist);
                                                            }
                                                            #endregion
                                                            #region VP slope
                                                            if (checkBox_slope_band.Checked == true && checkBox_delete_vp.Checked == false)
                                                            {
                                                                if (_AGEN_mainform.Vw_slope_height > 0 && _AGEN_mainform.Vw_ps_slope_y > 0)
                                                                {

                                                                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                                                                    {
                                                                        ms_point = new Point3d((double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_x], (double)_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_y], 0);
                                                                        double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_M1]);
                                                                        double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[si_index][_AGEN_mainform.Col_M2]);
                                                                        Point3d PSpoint_slope = new Point3d(_AGEN_mainform.Vw_ps_slope_x, _AGEN_mainform.Vw_ps_slope_y, 0);

                                                                        string Layer_name_prof_viewport_old = "VP_Slope";

                                                                        foreach (ObjectId odid in BtrecordPS)
                                                                        {
                                                                            Entity ent1 = Trans2.GetObject(odid, OpenMode.ForRead) as Entity;
                                                                            if (ent1 != null)
                                                                            {
                                                                                if (ent1.Layer.ToLower() == Layer_name_prof_viewport_old.ToLower())
                                                                                {
                                                                                    ent1.UpgradeOpen();
                                                                                    ent1.Erase();
                                                                                }
                                                                            }
                                                                        }

                                                                        double Hexag = 1;
                                                                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
                                                                        {
                                                                            Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
                                                                        }
                                                                        if (Hexag == 0) Hexag = 1;

                                                                        double Sta_at_0 = Functions.Station_equation_of(0, _AGEN_mainform.dt_station_equation);
                                                                        Creaza_viewport_slope(Trans2, Database2, BtrecordPS, Sta_at_0, Hexag, M1, M2, PSpoint_slope);
                                                                    }
                                                                }
                                                            }

                                                            #endregion

                                                            #region VP custom bands
                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0 && checkBox_delete_vp.Checked == false)
                                                            {
                                                                if (checkBox1.Visible == true)
                                                                {
                                                                    if (checkBox1.Checked == true)
                                                                    {

                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 1)
                                                                        {

                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ps_y"]);

                                                                                double y1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_y"]) -
                                                                                                           0.5 * height1 / scale_cust - si_index * _AGEN_mainform.Band_Separation / scale_cust;

                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["viewport_ms_x"]), y1, 0);

                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, scale_cust, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                if (checkBox2.Visible == true)
                                                                {
                                                                    if (checkBox2.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 2)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[1]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                if (checkBox3.Visible == true)
                                                                {
                                                                    if (checkBox3.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 3)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[2]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                if (checkBox4.Visible == true)
                                                                {
                                                                    if (checkBox4.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 4)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[3]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                if (checkBox5.Visible == true)
                                                                {
                                                                    if (checkBox5.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 5)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[4]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                                if (checkBox6.Visible == true)
                                                                {
                                                                    if (checkBox6.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 6)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[5]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox7.Visible == true)
                                                                {
                                                                    if (checkBox7.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 7)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[6]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox8.Visible == true)
                                                                {
                                                                    if (checkBox8.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 8)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[7]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox9.Visible == true)
                                                                {
                                                                    if (checkBox9.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 9)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[8]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox10.Visible == true)
                                                                {
                                                                    if (checkBox10.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 10)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[9]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox11.Visible == true)
                                                                {
                                                                    if (checkBox11.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 11)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[10]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox12.Visible == true)
                                                                {
                                                                    if (checkBox12.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 12)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[11]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox13.Visible == true)
                                                                {
                                                                    if (checkBox13.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 13)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[12]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox14.Visible == true)
                                                                {
                                                                    if (checkBox14.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 14)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[13]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox15.Visible == true)
                                                                {
                                                                    if (checkBox15.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 14)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[14]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }


                                                                if (checkBox16.Visible == true)
                                                                {
                                                                    if (checkBox16.Checked == true)
                                                                    {
                                                                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 16)
                                                                        {
                                                                            if (_AGEN_mainform.Data_Table_custom_bands.Rows[15]["band_name"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_height"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_width"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_y"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_x"] != DBNull.Value &&
                                                                                _AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_y"] != DBNull.Value)
                                                                            {
                                                                                string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["band_name"]) + "VP";

                                                                                height1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_height"]);
                                                                                width1 = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_width"]);

                                                                                double xps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_x"]);
                                                                                double yps = Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ps_y"]);


                                                                                ms_pt = new Point3d(Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_x"]),
                                                                                                    Convert.ToDouble(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["viewport_ms_y"]) - 0.5 * height1 -

                                                                                                   si_index * _AGEN_mainform.Band_Separation, 0);
                                                                                ps_pt = new Point3d(xps, yps, 0);

                                                                                Creaza_viewport_on_alignment_at_generation(Database2, LN, ms_pt, ps_pt, width1, height1, 1, 0);

                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                            #endregion

                                                            Trans2.Commit();
                                                        }
                                                        HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                        Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                                        #region profile band vp rectangles
                                                        if (checkBox_profile_band.Checked == true && checkBox_delete_vp.Checked == false)
                                                        {

                                                            Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                                            Polyline polyVP1 = new Polyline();

                                                            polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                            polyVP1.ColorIndex = 1;

                                                            polyVP1.AddVertexAt(0, new Point2d(ms_point.X - lr * (l_main / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                            polyVP1.AddVertexAt(1, new Point2d(ms_point.X - lr * (l_main / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                            polyVP1.AddVertexAt(2, new Point2d(ms_point.X + lr * (l_main / 2) / _AGEN_mainform.Vw_scale, ms_point.Y + (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                            polyVP1.AddVertexAt(3, new Point2d(ms_point.X + lr * (l_main / 2) / _AGEN_mainform.Vw_scale, ms_point.Y - (h_main / 2) / _AGEN_mainform.Vw_scale), 0, 0, 0);
                                                            polyVP1.Closed = true;

                                                            Btrecord.AppendEntity(polyVP1);
                                                            Trans1.AddNewlyCreatedDBObject(polyVP1, true);


                                                        }
                                                        #endregion


                                                        #region tblk band rectangles
                                                        if (checkBox_tblk.Checked == true && checkBox_delete_vp.Checked == false)
                                                        {
                                                            ms_pt = new Point3d(_AGEN_mainform.Point0_tblk.X, (_AGEN_mainform.Point0_tblk.Y - _AGEN_mainform.Vw_tblk_height / 2) - si_index * _AGEN_mainform.tblk_separation, 0);

                                                            if (_AGEN_mainform.tblk_twist == 90 * Math.PI / 180)
                                                            {
                                                                ms_pt = new Point3d(_AGEN_mainform.Point0_tblk.X, (_AGEN_mainform.Point0_tblk.Y - _AGEN_mainform.Vw_tblk_width / 2) - si_index * _AGEN_mainform.tblk_separation, 0);
                                                            }
                                                            if (_AGEN_mainform.tblk_twist == 0)
                                                            {
                                                                draw_custom_vp_rectangles(Trans1, Btrecord, lr, ms_pt, _AGEN_mainform.Vw_tblk_width, _AGEN_mainform.Vw_tblk_height, nume_fara_ext);
                                                            }
                                                            if (_AGEN_mainform.tblk_twist == Math.PI / 2)
                                                            {
                                                                draw_custom_vp_rectangles(Trans1, Btrecord, lr, ms_pt, _AGEN_mainform.Vw_tblk_height, _AGEN_mainform.Vw_tblk_width, nume_fara_ext);
                                                            }
                                                        }
                                                        #endregion


                                                        #region MULTIPLE profile vp rectangles
                                                        if (checkBox_mult_vp_prof.Checked == true && checkBox_delete_vp.Checked == false)
                                                        {
                                                            if (lista_poly.Count > 0)
                                                            {
                                                                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);

                                                                for (int m = 0; m < lista_poly.Count; ++m)
                                                                {

                                                                    Polyline polyVP1 = new Polyline();
                                                                    polyVP1 = lista_poly[m];
                                                                    polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                                    polyVP1.ColorIndex = 1;
                                                                    Btrecord.AppendEntity(polyVP1);
                                                                    Trans1.AddNewlyCreatedDBObject(polyVP1, true);
                                                                }
                                                            }
                                                        }
                                                        #endregion


                                                        #region nodata band vp rectangles
                                                        if (checkBox_no_data_band.Checked == true)
                                                        {
                                                            Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                                            Polyline polyVP1 = new Polyline();

                                                            polyVP1.Layer = _AGEN_mainform.layer_no_plot;
                                                            polyVP1.ColorIndex = 1;

                                                            polyVP1.AddVertexAt(0, new Point2d(x_nd_ms - (w_nd / 2) / scale_nd, y_no_data_c + 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                            polyVP1.AddVertexAt(1, new Point2d(x_nd_ms + (w_nd / 2) / scale_nd, y_no_data_c + 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                            polyVP1.AddVertexAt(2, new Point2d(x_nd_ms + (w_nd / 2) / scale_nd, y_no_data_c - 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                            polyVP1.AddVertexAt(3, new Point2d(x_nd_ms - (w_nd / 2) / scale_nd, y_no_data_c - 0.5 * h_nd / scale_nd), 0, 0, 0);
                                                            polyVP1.Closed = true;

                                                            Btrecord.AppendEntity(polyVP1);
                                                            Trans1.AddNewlyCreatedDBObject(polyVP1, true);

                                                        }
                                                        #endregion


                                                    }
                                                }
                                            }
                                        }
                                    }
                                    MessageBox.Show("done");
                                }
                                #endregion


                                Trans1.Commit();
                            }
                        }
                    }
                }

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
            catch (System.AccessViolationException ex1)
            {
                MessageBox.Show(ex1.Message);

            }


            _AGEN_mainform.tpage_processing.Hide();
            Freeze_operations = false;

        }


        private void Creaza_viewport_on_alignment_at_generation(Database Database2, string Layer_name_Viewport, Point3d MSpoint, Point3d PSpoint, double width1, double height1, double scale1, double twist1)
        {
            HostApplicationServices.WorkingDatabase = Database2;
            Functions.Creaza_layer_on_database(Database2, Layer_name_Viewport, 4, false);
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {
                Functions.make_first_layout_active(Trans2, Database2);
                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                BtrecordPS.UpgradeOpen();
                //Layout Layout1 = Functions.get_first_layout(Trans2, Database2);
                //Layout1.UpgradeOpen();
                //Layout1.LayoutName = nume1;
                foreach (ObjectId id1 in BtrecordPS)
                {
                    Viewport Viewport_old = Trans2.GetObject(id1, OpenMode.ForRead) as Viewport;
                    if (Viewport_old != null)
                    {
                        if (Viewport_old.Layer == Layer_name_Viewport)
                        {
                            Viewport_old.UpgradeOpen();
                            Viewport_old.Erase();
                        }
                    }
                }


                Viewport new_viewport = Functions.Create_viewport(MSpoint, PSpoint, width1, height1, scale1, twist1);
                new_viewport.Layer = Layer_name_Viewport;
                BtrecordPS.AppendEntity(new_viewport);
                Trans2.AddNewlyCreatedDBObject(new_viewport, true);

                Trans2.Commit();
            }
        }



        private void Creaza_viewport_on_alignment_on_existing(Database Database2, string Layer_name_Viewport, Point3d MSpoint, Point3d PSpoint, double width1, double height1, double scale1, double twist1)
        {
            HostApplicationServices.WorkingDatabase = Database2;
            Functions.Creaza_layer_on_database(Database2, Layer_name_Viewport, 4, false);

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {
                Functions.creaza_anno_scales(Database2);
                var ocm = Database2.ObjectContextManager;
                var occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");


                Functions.make_first_layout_active(Trans2, Database2);
                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                BtrecordPS.UpgradeOpen();

                List<string> lista_del = new List<string>();
                lista_del.Add(Layer_name_Viewport);

                foreach (ObjectId id1 in BtrecordPS)
                {
                    Viewport Viewport_old = Trans2.GetObject(id1, OpenMode.ForRead) as Viewport;

                    if (Viewport_old != null)
                    {
                        if (lista_del.Contains(Viewport_old.Layer) == true)
                        {
                            Viewport_old.UpgradeOpen();
                            Viewport_old.Erase();
                        }
                    }
                }

                Viewport new_viewport = Functions.Create_viewport(MSpoint, PSpoint, width1, height1, scale1, twist1);

                new_viewport.Layer = Layer_name_Viewport;
                BtrecordPS.AppendEntity(new_viewport);



                #region annotation implementation
                string anno_name = "xxx";
                if (Math.Round(scale1, 1) == 0.1)
                {
                    anno_name = "_1:10";
                }
                if (Math.Round(scale1, 2) == 0.05)
                {
                    anno_name = "_1:20";
                }
                if (Math.Round(scale1, 3) == 0.033)
                {
                    anno_name = "_1:30";
                }
                if (Math.Round(scale1, 3) == 0.025)
                {
                    anno_name = "_1:40";
                }
                if (Math.Round(scale1, 2) == 0.02)
                {
                    anno_name = "_1:50";
                }
                if (Math.Round(scale1, 3) == 0.017)
                {
                    anno_name = "_1:60";
                }
                if (Math.Round(scale1, 2) == 0.01)
                {
                    anno_name = "_1:100";
                }
                if (Math.Round(scale1, 3) == 0.005)
                {
                    anno_name = "_1:200";
                }
                if (Math.Round(scale1, 4) == 0.0033)
                {
                    anno_name = "_1:300";
                }
                if (Math.Round(scale1, 4) == 0.0025)
                {
                    anno_name = "_1:400";
                }
                if (Math.Round(scale1, 3) == 0.002)
                {
                    anno_name = "_1:500";
                }
                if (Math.Round(scale1, 4) == 0.0017)
                {
                    anno_name = "_1:600";
                }

                foreach (var context1 in occ)
                {
                    if (context1.Name == anno_name)
                    {
                        new_viewport.AnnotationScale = (AnnotationScale)context1;
                    }
                }
                #endregion

                Trans2.AddNewlyCreatedDBObject(new_viewport, true);

                ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                oBJiD_COL.Add(new_viewport.ObjectId);
                DrawOrderTable DrawOrderTable2 = Trans2.GetObject(BtrecordPS.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                DrawOrderTable2.MoveToBottom(oBJiD_COL);




                Trans2.Commit();
            }
        }

        private void Delete_viewport_on_existing_alignment(Database Database2)
        {
            HostApplicationServices.WorkingDatabase = Database2;

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {
                Functions.make_first_layout_active(Trans2, Database2);
                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                BtrecordPS.UpgradeOpen();

                List<string> lista_del = new List<string>();
                if (checkBox_plan_view.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_Main_Viewport);
                if (checkBox_materials.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_material_Viewport);
                if (checkBox_crossing.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_crossing_Viewport);
                if (checkBox_ownership.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_ownership_Viewport);
                if (checkBox_profile.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_prof_main_viewport);
                if (checkBox_profile.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_prof_side_viewport);
                if (checkBox_profile_band.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_profband_Viewport);
                if (checkBox_tblk.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_tblk_Viewport);
                if (checkBox_no_data_band.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_no_data_band_Viewport);

                if (checkBox_slope_band.Checked == true) lista_del.Add("VP_Slope");

                if (checkBoxx1.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_extra1_Viewport);
                if (checkBoxx2.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_extra2_Viewport);
                if (checkBoxx3.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_extra3_Viewport);
                if (checkBoxx4.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_extra4_Viewport);
                if (checkBoxx5.Checked == true) lista_del.Add(_AGEN_mainform.Layer_name_extra5_Viewport);

                #region CUSTOM BANDS VP
                if (checkBox1.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 1)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[0]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox2.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 2)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[1]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox3.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 3)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[2]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox4.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 4)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[3]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox5.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 5)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[4]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox6.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 6)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[5]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox7.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 7)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[6]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox8.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 8)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[7]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox9.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 9)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[8]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox10.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 10)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[9]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox11.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 11)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[10]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox12.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 12)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[11]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox13.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 13)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[12]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox14.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 14)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[13]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox15.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 15)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[14]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                if (checkBox16.Checked == true)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null && _AGEN_mainform.Data_Table_custom_bands.Rows.Count >= 16)
                    {
                        string LN = "AGEN_" + Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[15]["band_name"]) + "VP";
                        lista_del.Add(LN);
                    }
                }

                #endregion

                foreach (ObjectId id1 in BtrecordPS)
                {
                    Viewport Viewport_old = Trans2.GetObject(id1, OpenMode.ForRead) as Viewport;

                    if (Viewport_old != null)
                    {
                        if (lista_del.Contains(Viewport_old.Layer) == true)
                        {
                            Viewport_old.UpgradeOpen();
                            Viewport_old.Erase();
                        }
                    }
                    if (checkBox_profile.Checked == true)
                    {
                        Entity ent1 = Trans2.GetObject(id1, OpenMode.ForRead) as Entity;
                        if (ent1 != null)
                        {
                            if (lista_del.Contains(ent1.Layer) == true)
                            {
                                ent1.UpgradeOpen();
                                ent1.Erase();
                            }
                        }
                    }

                }

                Trans2.Commit();
            }
        }

        private System.Data.DataTable create_profile_poly_definition(string File1)
        {
            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the profile data file does not exist");
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);

                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                if (segment1 == "not defined") segment1 = "";

                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                {
                    if (wsh1.Name == "pdc1_" + segment1)
                    {
                        W2 = wsh1;
                    }
                    if (wsh1.Name == "pdc2_" + segment1)
                    {
                        W1 = wsh1;
                    }
                }
                if (W1 == null || W2 == null)
                {
                    MessageBox.Show("No profile config defined");
                    return null;
                }
                try
                {
                    string s1 = Convert.ToString(W1.Range["B1"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_texth = Convert.ToDouble(s1);
                    }

                    s1 = Convert.ToString(W1.Range["B2"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_x0 = Convert.ToDouble(s1);
                    }

                    s1 = Convert.ToString(W1.Range["B3"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_y0 = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B4"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_x_left = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B5"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_x_right = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B6"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_y_down = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B7"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_hexag = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B8"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_vexag = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B9"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_down_el = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B10"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_up_el = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B11"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_start_sta = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B12"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_end_sta = Convert.ToDouble(s1);
                    }
                    s1 = Convert.ToString(W1.Range["B13"].Value2);
                    if (Functions.IsNumeric(s1) == true)
                    {
                        _AGEN_mainform.prof_width_lr = Convert.ToDouble(s1);
                    }
                    dt2 = Functions.Build_Data_table_prof_poly_from_excel(W2, _AGEN_mainform.Start_row_1 + 1);

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
            return dt2;
        }



        private void Creaza_viewports_profile(string file1, Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS, System.Data.DataTable Dtpoly, double M1, double M2, Point3d PSpoint, int index_Data_Table_prof_dwg)
        {
            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;

            string nume1 = System.IO.Path.GetFileNameWithoutExtension(file1);

            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_prof_main_viewport, 7, true);
            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_prof_side_viewport, 7, false);

            foreach (ObjectId id1 in BtrecordPS)
            {
                Entity Ent1 = Trans2.GetObject(id1, OpenMode.ForRead) as Entity;
                if (Ent1 != null)
                {
                    if (Ent1.Layer == _AGEN_mainform.Layer_name_prof_main_viewport || Ent1.Layer == _AGEN_mainform.Layer_name_prof_side_viewport)
                    {
                        Ent1.UpgradeOpen();
                        Ent1.Erase();
                    }
                }
            }

            Polyline poly_g = new Polyline();
            int idx0 = 0;
            for (int m = 0; m < Dtpoly.Rows.Count; ++m)
            {
                if (Dtpoly.Rows[m][_AGEN_mainform.Col_x] != DBNull.Value && Dtpoly.Rows[m][_AGEN_mainform.Col_y] != DBNull.Value)
                {
                    double x11 = Convert.ToDouble(Dtpoly.Rows[m][_AGEN_mainform.Col_x]);
                    double y11 = Convert.ToDouble(Dtpoly.Rows[m][_AGEN_mainform.Col_y]);
                    poly_g.AddVertexAt(idx0, new Point2d(x11, y11), 0, 0, 0);
                    idx0 = idx0 + 1;
                }
            }

            double x0 = _AGEN_mainform.prof_x0;
            double y0 = _AGEN_mainform.prof_y0;
            Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(x0 + lr * (M1 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag, y0, 0), new Point3d(x0 + lr * (M1 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag, y0 + 10000000, 0));
            Autodesk.AutoCAD.DatabaseServices.Line Line2 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(x0 + lr * (M2 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag, y0, 0), new Point3d(x0 + lr * (M2 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag, y0 + 10000000, 0));
            Point3dCollection Col1 = Functions.Intersect_on_both_operands(Line1, poly_g);
            Point3dCollection Col2 = Functions.Intersect_on_both_operands(Line2, poly_g);
            Point3d Pt1 = new Point3d(0, 0, 0);
            if (Col1.Count == 1)
            {
                Pt1 = Col1[0];
            }
            Point3d Pt2 = new Point3d(0, 0, 0);
            if (Col2.Count == 1)
            {
                Pt2 = Col2[0];
            }
            double x1 = Pt1.X;
            double y1 = Pt1.Y;
            double x2 = Pt2.X;
            double y2 = Pt2.Y;

            if (x1 != 0 || y1 != 0 || x2 != 0 || y2 != 0)
            {
                if (x1 == 0 && y1 == 0 && (x2 != 0 || y2 != 0))
                {
                    x1 = x0 + lr * (M1 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag;
                    y1 = y2;
                }

                if (x2 == 0 && y2 == 0 && (x1 != 0 || y1 != 0))
                {
                    x2 = x0 + lr * (M2 - _AGEN_mainform.prof_start_sta) * _AGEN_mainform.prof_hexag;
                    y2 = y1;
                }

                Point3d MSpoint = new Point3d((x1 + x2) / 2, (y1 + y2) / 2, 0);
                double width1 = lr * (x2 - x1) * _AGEN_mainform.Vw_scale;
                string prefix_sta = "STA. ";
                double a1 = 3 * _AGEN_mainform.Vw_scale;
                double height1 = _AGEN_mainform.Vw_prof_height - a1 * _AGEN_mainform.prof_texth;
                PSpoint = new Point3d(PSpoint.X, PSpoint.Y + a1 * _AGEN_mainform.prof_texth / 2, 0);
                double prof_width_lr_ps = _AGEN_mainform.prof_width_lr * _AGEN_mainform.Vw_scale;

                Viewport Viewport_prof_main = Functions.Create_viewport(MSpoint, PSpoint, width1, height1, _AGEN_mainform.Vw_scale, 0);
                Viewport_prof_main.Layer = _AGEN_mainform.Layer_name_prof_main_viewport;
                BtrecordPS.AppendEntity(Viewport_prof_main);
                Trans2.AddNewlyCreatedDBObject(Viewport_prof_main, true);

                Viewport Viewport_left = Functions.Create_viewport(new Point3d(_AGEN_mainform.prof_x_left, MSpoint.Y, 0), new Point3d(PSpoint.X - width1 / 2 - a1 * _AGEN_mainform.prof_texth - prof_width_lr_ps / 2, PSpoint.Y, 0), prof_width_lr_ps, height1, _AGEN_mainform.Vw_scale, 0);
                Viewport_left.Layer = _AGEN_mainform.Layer_name_prof_side_viewport;
                BtrecordPS.AppendEntity(Viewport_left);
                Trans2.AddNewlyCreatedDBObject(Viewport_left, true);

                Viewport Viewport_right = Functions.Create_viewport(new Point3d(_AGEN_mainform.prof_x_right, MSpoint.Y, 0), new Point3d(PSpoint.X + width1 / 2 + a1 * _AGEN_mainform.prof_texth + prof_width_lr_ps / 2, PSpoint.Y, 0), prof_width_lr_ps, height1, _AGEN_mainform.Vw_scale, 0);
                Viewport_right.Layer = _AGEN_mainform.Layer_name_prof_side_viewport;
                BtrecordPS.AppendEntity(Viewport_right);
                Trans2.AddNewlyCreatedDBObject(Viewport_right, true);

                Viewport Viewport_down = Functions.Create_viewport(new Point3d(MSpoint.X, _AGEN_mainform.prof_y_down, 0), new Point3d(PSpoint.X, PSpoint.Y - (height1 + a1 * _AGEN_mainform.prof_texth) / 2, 0), width1 + 4 * _AGEN_mainform.prof_texth * _AGEN_mainform.Vw_scale + 2 * prof_width_lr_ps, a1 * _AGEN_mainform.prof_texth, _AGEN_mainform.Vw_scale, 0);
                Viewport_down.Layer = _AGEN_mainform.Layer_name_prof_side_viewport;
                BtrecordPS.AppendEntity(Viewport_down);
                Trans2.AddNewlyCreatedDBObject(Viewport_down, true);

                MText mtext_label_left = new MText();
                mtext_label_left.Contents = prefix_sta + Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(M1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                mtext_label_left.TextHeight = _AGEN_mainform.prof_texth * _AGEN_mainform.Vw_scale;
                mtext_label_left.Rotation = Math.PI / 2;
                mtext_label_left.Attachment = AttachmentPoint.MiddleCenter;
                mtext_label_left.Location = new Point3d(PSpoint.X - lr * width1 / 2 - lr * (a1 / 2) * _AGEN_mainform.prof_texth, PSpoint.Y, 0);
                mtext_label_left.Layer = _AGEN_mainform.Layer_name_prof_main_viewport;
                BtrecordPS.AppendEntity(mtext_label_left);
                Trans2.AddNewlyCreatedDBObject(mtext_label_left, true);

                MText mtext_label_right = new MText();
                mtext_label_right.Contents = prefix_sta + Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                mtext_label_right.TextHeight = _AGEN_mainform.prof_texth * _AGEN_mainform.Vw_scale;
                mtext_label_right.Rotation = Math.PI / 2;
                mtext_label_right.Attachment = AttachmentPoint.MiddleCenter;
                mtext_label_right.Location = new Point3d(PSpoint.X + lr * width1 / 2 + lr * (a1 / 2) * _AGEN_mainform.prof_texth, PSpoint.Y, 0);
                mtext_label_right.Layer = _AGEN_mainform.Layer_name_prof_main_viewport;
                BtrecordPS.AppendEntity(mtext_label_right);
                Trans2.AddNewlyCreatedDBObject(mtext_label_right, true);

            }
        }



        private void Creaza_viewport_slope(Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS, double Sta_at_0, double Hexag, double M1, double M2, Point3d PSpoint)
        {






            Functions.Creaza_layer_on_database(Database2, "VP_Slope", 7, false);



            double x0 = _AGEN_mainform.Point0_slope.X;
            double y0 = _AGEN_mainform.Point0_slope.Y;



            Point3d MSpoint = new Point3d(x0 + Hexag * ((M1 + M2) / 2 - Sta_at_0), y0 + 0.5 * _AGEN_mainform.Vw_slope_height / _AGEN_mainform.Vw_scale, 0);


            double width1 = Hexag * (M2 - M1) * _AGEN_mainform.Vw_scale;


            Viewport Viewport_slope = Functions.Create_viewport(MSpoint, PSpoint, width1, _AGEN_mainform.Vw_slope_height, _AGEN_mainform.Vw_scale, 0);
            Viewport_slope.Layer = "VP_Slope";
            BtrecordPS.AppendEntity(Viewport_slope);
            Trans2.AddNewlyCreatedDBObject(Viewport_slope, true);





        }


        private void Creaza_viewports_profile_band(Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS,
                                                    Point3d MSpoint_main, Point3d PSpoint_main,
                                                    double width1, double height_main, string Layer_name_prof_viewport)
        {


            Functions.Creaza_layer_on_database(Database2, Layer_name_prof_viewport, 7, false);

            Viewport Viewport_prof_main = Functions.Create_viewport(MSpoint_main, PSpoint_main, width1, height_main, _AGEN_mainform.Vw_scale, 0);
            Viewport_prof_main.Layer = Layer_name_prof_viewport;
            BtrecordPS.AppendEntity(Viewport_prof_main);
            Trans2.AddNewlyCreatedDBObject(Viewport_prof_main, true);

        }



        private void Creaza_viewports_no_data_band(Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS,
                                            Point3d MSpoint_no_data, Point3d PSpoint_no_data,
                                            double width1, double h1, double scale1, string layer1)
        {


            Functions.Creaza_layer_on_database(Database2, layer1, 7, false);

            Viewport Viewport_prof_no_data = Functions.Create_viewport(MSpoint_no_data, PSpoint_no_data, width1, h1, scale1, 0);
            Viewport_prof_no_data.Layer = layer1;
            BtrecordPS.AppendEntity(Viewport_prof_no_data);
            Trans2.AddNewlyCreatedDBObject(Viewport_prof_no_data, true);

        }


        private void Creaza_viewports_property_at_alignments_generation(Autodesk.AutoCAD.DatabaseServices.Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS, Point3d MSpoint, int Sheet_index_i)
        {
            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_ownership_Viewport, 4, false);
            Viewport Viewport_property = Functions.Create_viewport(MSpoint, new Point3d(_AGEN_mainform.Vw_ps_prop_x, _AGEN_mainform.Vw_ps_prop_y, 0), _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_prop_height, 1, 0);
            Viewport_property.Layer = _AGEN_mainform.Layer_name_ownership_Viewport;
            BtrecordPS.AppendEntity(Viewport_property);
            Trans2.AddNewlyCreatedDBObject(Viewport_property, true);

        }

        private void Creaza_viewports_crossing_at_alignments_generation(Autodesk.AutoCAD.DatabaseServices.Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS, Point3d MSpoint, int Sheet_index_i)
        {
            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_crossing_Viewport, 4, false);
            Viewport Viewport_crossing = Functions.Create_viewport(MSpoint, new Point3d(_AGEN_mainform.Vw_ps_cross_x, _AGEN_mainform.Vw_ps_cross_y, 0), _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_cross_height, 1, 0);
            Viewport_crossing.Layer = _AGEN_mainform.Layer_name_crossing_Viewport;
            BtrecordPS.AppendEntity(Viewport_crossing);
            Trans2.AddNewlyCreatedDBObject(Viewport_crossing, true);

        }

        private void Creaza_viewports_material_at_alignments_generation(Autodesk.AutoCAD.DatabaseServices.Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS, Point3d MSpoint, int Sheet_index_i)
        {
            Functions.Creaza_layer_on_database(Database2, _AGEN_mainform.Layer_name_material_Viewport, 4, false);
            Viewport Viewport_material = Functions.Create_viewport(MSpoint, new Point3d(_AGEN_mainform.Vw_ps_mat_x, _AGEN_mainform.Vw_ps_mat_y, 0), _AGEN_mainform.Vw_width, _AGEN_mainform.Vw_mat_height, 1, 0);
            Viewport_material.Layer = _AGEN_mainform.Layer_name_material_Viewport;
            BtrecordPS.AppendEntity(Viewport_material);
            Trans2.AddNewlyCreatedDBObject(Viewport_material, true);

        }

        private System.Data.DataTable Load_existing_profile_graph(String File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the profile data file does not exist");
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {


                    dt2 = Functions.Build_Data_table_profile_from_excel(W1, _AGEN_mainform.Start_row_graph_profile + 1);




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
            return dt2;

        }



        private void dataGridView_align_created_Click(object sender, EventArgs e)
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

        private void dataGridView_align_created_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_align_created.CurrentCell = dataGridView_align_created.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_open_alignment.Show(Cursor.Position);
                ContextMenuStrip_open_alignment.Visible = true;
            }
            else
            {
                ContextMenuStrip_open_alignment.Visible = false;
            }
        }

        public void Hide_panel_extra_bands()
        {
            panel_extra_view.Visible = false;
        }

        public void show_panel_extra_bands()
        {
            panel_extra_view.Visible = true;
        }

        public void Hide_checkBox_extra1()
        {
            checkBoxx1.Visible = false;
        }

        public void show_checkBox_extra1()
        {
            checkBoxx1.Visible = true;
        }

        public void Hide_checkBox_extra2()
        {
            checkBoxx2.Visible = false;
        }

        public void show_checkBox_extra2()
        {
            checkBoxx2.Visible = true;
        }

        public void Hide_checkBox_extra3()
        {
            checkBoxx3.Visible = false;
        }

        public void show_checkBox_extra3()
        {
            checkBoxx3.Visible = true;
        }
        public void Hide_checkBox_extra4()
        {
            checkBoxx4.Visible = false;
        }

        public void show_checkBox_extra4()
        {
            checkBoxx4.Visible = true;
        }
        public void Hide_checkBox_extra5()
        {
            checkBoxx5.Visible = false;
        }

        public void show_checkBox_extra5()
        {
            checkBoxx5.Visible = true;
        }


        public void Hide_panel_custom_bands()
        {
            panel_custom_bands.Visible = false;
        }

        public void show_panel_custom_bands()
        {
            panel_custom_bands.Visible = true;
        }

        public void Hide_checkBox1()
        {
            checkBox1.Visible = false;
        }

        public void show_checkBox1()
        {
            checkBox1.Visible = true;
        }

        public void Hide_checkBox2()
        {
            checkBox2.Visible = false;
        }

        public void show_checkBox2()
        {
            checkBox2.Visible = true;
        }

        public void Hide_checkBox3()
        {
            checkBox3.Visible = false;
        }

        public void show_checkBox3()
        {
            checkBox3.Visible = true;
        }

        public void Hide_checkBox4()
        {
            checkBox4.Visible = false;
        }

        public void show_checkBox4()
        {
            checkBox4.Visible = true;
        }
        public void Hide_checkBox5()
        {
            checkBox5.Visible = false;
        }

        public void show_checkBox5()
        {
            checkBox5.Visible = true;
        }
        public void Hide_checkBox6()
        {
            checkBox6.Visible = false;
        }

        public void show_checkBox6()
        {
            checkBox6.Visible = true;
        }
        public void Hide_checkBox7()
        {
            checkBox7.Visible = false;
        }

        public void show_checkBox7()
        {
            checkBox7.Visible = true;
        }
        public void Hide_checkBox8()
        {
            checkBox8.Visible = false;
        }

        public void show_checkBox8()
        {
            checkBox8.Visible = true;
        }
        public void Hide_checkBox9()
        {
            checkBox9.Visible = false;
        }

        public void show_checkBox9()
        {
            checkBox9.Visible = true;
        }
        public void Hide_checkBox10()
        {
            checkBox10.Visible = false;
        }

        public void show_checkBox10()
        {
            checkBox10.Visible = true;
        }
        public void Hide_checkBox11()
        {
            checkBox11.Visible = false;
        }

        public void show_checkBox11()
        {
            checkBox11.Visible = true;
        }
        public void Hide_checkBox12()
        {
            checkBox12.Visible = false;
        }

        public void show_checkBox12()
        {
            checkBox12.Visible = true;
        }
        public void Hide_checkBox13()
        {
            checkBox13.Visible = false;
        }

        public void show_checkBox13()
        {
            checkBox13.Visible = true;
        }
        public void Hide_checkBox14()
        {
            checkBox14.Visible = false;
        }

        public void show_checkBox14()
        {
            checkBox14.Visible = true;
        }
        public void Hide_checkBox15()
        {
            checkBox15.Visible = false;
        }

        public void show_checkBox15()
        {
            checkBox15.Visible = true;
        }
        public void Hide_checkBox16()
        {
            checkBox16.Visible = false;
        }

        public void show_checkBox16()
        {
            checkBox16.Visible = true;
        }

        public void set_txt_checkBox1(string txt)
        {
            checkBox1.Text = txt;
        }
        public void set_txt_checkBox2(string txt)
        {
            checkBox2.Text = txt;
        }
        public void set_txt_checkBox3(string txt)
        {
            checkBox3.Text = txt;
        }
        public void set_txt_checkBox4(string txt)
        {
            checkBox4.Text = txt;
        }
        public void set_txt_checkBox5(string txt)
        {
            checkBox5.Text = txt;
        }
        public void set_txt_checkBox6(string txt)
        {
            checkBox6.Text = txt;
        }
        public void set_txt_checkBox7(string txt)
        {
            checkBox7.Text = txt;
        }
        public void set_txt_checkBox8(string txt)
        {
            checkBox8.Text = txt;
        }
        public void set_txt_checkBox9(string txt)
        {
            checkBox9.Text = txt;
        }
        public void set_txt_checkBox10(string txt)
        {
            checkBox10.Text = txt;
        }
        public void set_txt_checkBox11(string txt)
        {
            checkBox11.Text = txt;
        }
        public void set_txt_checkBox12(string txt)
        {
            checkBox12.Text = txt;
        }
        public void set_txt_checkBox13(string txt)
        {
            checkBox13.Text = txt;
        }
        public void set_txt_checkBox14(string txt)
        {
            checkBox14.Text = txt;
        }
        public void set_txt_checkBox15(string txt)
        {
            checkBox15.Text = txt;
        }
        public void set_txt_checkBox16(string txt)
        {
            checkBox16.Text = txt;
        }

        private void delete_2nd_layout(object sender, EventArgs e)
        {



            if (Display_dt != null)
            {
                if (Display_dt.Rows.Count > 0)
                {


                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            for (int i = 0; i < Display_dt.Rows.Count; ++i)
                            {
                                if (Display_dt.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value)
                                {
                                    string file1 = Display_dt.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                    string nume_fara_ext = System.IO.Path.GetFileNameWithoutExtension(file1);

                                    if (System.IO.File.Exists(file1) == true)
                                    {
                                        using (Database Database2 = new Database(false, true))
                                        {
                                            Database2.ReadDwgFile(file1, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                            //System.IO.FileShare.ReadWrite, false, null);
                                            Database2.CloseInput(true);

                                            HostApplicationServices.WorkingDatabase = Database2;


                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                            {



                                                DBDictionary Layoutdict = (DBDictionary)Trans2.GetObject(Database2.LayoutDictionaryId, OpenMode.ForWrite);

                                                LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;


                                                foreach (DBDictionaryEntry entry in Layoutdict)
                                                {
                                                    Layout Layout0 = (Layout)Trans2.GetObject(LayoutManager1.GetLayoutId(entry.Key), OpenMode.ForWrite);
                                                    if (Layout0.TabOrder == 2)
                                                    {

                                                        Layout0.Erase();
                                                    }

                                                }












                                                Trans2.Commit();







                                            }




                                            HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                            Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                        }
                                    }
                                }
                            }

                            MessageBox.Show("done");

                        }
                    }
                }

            }
        }

        public void Delete_vp_from_existing_dwgs(String Layername, Database Database2, Point3d oldVP_PS, double h, double w)
        {
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {

                Functions.make_first_layout_active(Trans2, Database2);

                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                BtrecordPS.UpgradeOpen();

                double buffer1 = 5;

                foreach (ObjectId odid in BtrecordPS)
                {
                    Viewport Vp_old = Trans2.GetObject(odid, OpenMode.ForRead) as Viewport;
                    if (Vp_old != null)
                    {
                        if (Vp_old.Layer.ToLower() == Layername.ToLower() &&
                            Vp_old.CenterPoint.X < oldVP_PS.X + buffer1 && Vp_old.CenterPoint.X > oldVP_PS.X - buffer1 &&
                            Vp_old.CenterPoint.Y < oldVP_PS.Y + buffer1 && Vp_old.CenterPoint.Y > oldVP_PS.Y - buffer1 &&
                            Vp_old.Width < w + buffer1 && Vp_old.Width > w - buffer1 &&
                            Vp_old.Height < h + buffer1 && Vp_old.Height > h - buffer1)
                        {
                            Vp_old.UpgradeOpen();
                            Vp_old.Erase();
                        }
                    }

                }
                Trans2.Commit();

            }
        }

        public void Delete_all_vp_from_existing_dwgs(Transaction Trans2, Database Database2, BlockTableRecord BtrecordPS)
        {

            foreach (ObjectId odid in BtrecordPS)
            {
                Viewport Vp_old = Trans2.GetObject(odid, OpenMode.ForWrite) as Viewport;
                if (Vp_old != null)
                {
                    Vp_old.Erase();
                }

            }
        }
        public void Delete_profile_vp_from_existing_dwgs(Transaction Trans2, Database Database2, Point3d corner1, Point3d corner2)
        {



            BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
            BtrecordPS.UpgradeOpen();


            foreach (ObjectId odid in BtrecordPS)
            {
                Viewport Vp_old = Trans2.GetObject(odid, OpenMode.ForRead) as Viewport;
                if (Vp_old != null)
                {
                    if (
                        Vp_old.CenterPoint.X > corner1.X && Vp_old.CenterPoint.X < corner2.X &&
                        Vp_old.CenterPoint.Y > corner1.Y && Vp_old.CenterPoint.Y < corner2.Y
                        )
                    {
                        Vp_old.UpgradeOpen();
                        Vp_old.Erase();
                    }
                }

                MText mt1 = Trans2.GetObject(odid, OpenMode.ForRead) as MText;
                if (mt1 != null)
                {
                    if (mt1.Contents.Contains("STA.") == true)
                    {
                        if (
                        mt1.Location.X > corner1.X && mt1.Location.X < corner2.X &&
                        mt1.Location.Y > corner1.Y && mt1.Location.Y < corner2.Y
                        )
                        {
                            mt1.UpgradeOpen();
                            mt1.Erase();
                        }
                    }
                }

            }

        }

        public void Delete_blocks_from_existing_dwgs(List<string> lista_lower_names, Database Database2)
        {

            if (lista_lower_names.Count > 0)
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                {

                    Functions.make_first_layout_active(Trans2, Database2);

                    BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                    BtrecordPS.UpgradeOpen();


                    foreach (ObjectId odid in BtrecordPS)
                    {
                        BlockReference bl1 = Trans2.GetObject(odid, OpenMode.ForRead) as BlockReference;
                        if (bl1 != null)
                        {
                            string nume_block = "";

                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)Database2.BlockTableId.GetObject(OpenMode.ForRead);
                            BlockTableRecord Btr = null;
                            if (bl1.IsDynamicBlock == true)
                            {

                                Btr = (BlockTableRecord)Trans2.GetObject(bl1.DynamicBlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                nume_block = Btr.Name;
                            }
                            else
                            {
                                Btr = (BlockTableRecord)Trans2.GetObject(bl1.BlockTableRecord, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                nume_block = Btr.Name;
                            }

                            if (lista_lower_names.Contains(nume_block.ToLower()) == true)
                            {
                                bl1.UpgradeOpen();
                                bl1.Erase();
                            }

                        }

                    }
                    Trans2.Commit();

                }
            }
        }

        private void label_create_alignments_Click(object sender, EventArgs e)
        {

            if (panel_dan.Visible == false)
            {
                panel_dan.Visible = true;
            }
            else
            {
                panel_dan.Visible = false;
            }

        }

        private void button_load_dwgs_in_comboboxs1_Click(object sender, EventArgs e)
        {
            comboBox_start.Items.Clear();
            comboBox_start.Items.Add("");


            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
            if (segment1 == "not defined") segment1 = "";

            if (_AGEN_mainform.current_segment.ToLower() != segment1.ToLower())
            {
                _AGEN_mainform.tpage_setup.Build_sheet_index_dt_from_excel();
            }

            if (_AGEN_mainform.dt_sheet_index != null && _AGEN_mainform.dt_sheet_index.Rows.Count > 0)
            {
                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"] != DBNull.Value)
                    {
                        comboBox_start.Items.Add(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"]));

                    }
                }
            }
        }

        private void button_load_dwgs_in_comboboxs2_Click(object sender, EventArgs e)
        {
            comboBox_end.Items.Clear();
            comboBox_end.Items.Add("");

            string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
            if (segment1 == "not defined") segment1 = "";

            if (_AGEN_mainform.current_segment.ToLower() != segment1.ToLower())
            {
                _AGEN_mainform.tpage_setup.Build_sheet_index_dt_from_excel();
            }

            if (_AGEN_mainform.dt_sheet_index != null && _AGEN_mainform.dt_sheet_index.Rows.Count > 0)
            {
                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"] != DBNull.Value)
                    {
                        comboBox_end.Items.Add(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"]));
                    }
                }
            }
        }

        public static System.Data.DataTable Creaza_dt_image_datatable_structure()
        {

            System.Data.DataTable dt1 = new System.Data.DataTable();
            dt1.Columns.Add("Drawing_File", typeof(string));
            dt1.Columns.Add("Image_File", typeof(string));
            dt1.Columns.Add("World_File", typeof(string));
            dt1.Columns.Add("Width", typeof(double));
            dt1.Columns.Add("Height", typeof(double));
            dt1.Columns.Add("X", typeof(double));
            dt1.Columns.Add("Y", typeof(double));
            return dt1;
        }

        private System.Data.DataTable Build_Data_table_imagery_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            dt_image = Creaza_dt_image_datatable_structure();

            string Col1 = "B";
            string Col2 = "C";
            string Col3 = "D";

            Range range1 = W1.Range[Col1 + Start_row.ToString() + ":" + Col1 + "30000"];
            object[,] values1 = new object[30000, 1];
            values1 = range1.Value2;

            Range range2 = W1.Range[Col2 + Start_row.ToString() + ":" + Col2 + "30000"];
            object[,] values2 = new object[30000, 1];
            values2 = range2.Value2;

            Range range3 = W1.Range[Col3 + Start_row.ToString() + ":" + Col3 + "30000"];
            object[,] values3 = new object[30000, 1];
            values3 = range3.Value2;

            bool is_data = false;
            for (int i = 1; i <= values1.Length; ++i)
            {
                object Valoare1 = values1[i, 1];
                object Valoare2 = values2[i, 1];
                object Valoare3 = values3[i, 1];
                if (Valoare1 != null && Valoare2 != null && Valoare3 != null)
                {
                    dt_image.Rows.Add();
                    is_data = true;
                }
                else
                {
                    i = values1.Length + 1;
                }
            }


            if (is_data == false)
            {
                return dt_image;
            }

            int NrR = dt_image.Rows.Count;


            Microsoft.Office.Interop.Excel.Range range_val = W1.Range[W1.Cells[Start_row, 2], W1.Cells[NrR + Start_row - 1, 4]];

            object[,] values = new object[NrR - 1, 3];

            values = range_val.Value2;

            for (int i = 0; i < dt_image.Rows.Count; ++i)
            {
                for (int j = 0; j < 3; ++j)
                {
                    object Valoare = values[i + 1, j + 1];
                    if (Valoare == null) Valoare = DBNull.Value;
                    dt_image.Rows[i][j] = Valoare;
                }
            }




            return dt_image;

        }


        private void Button_output_location_Click(object sender, EventArgs e)
        {
            try
            {
                string Output_folder = _AGEN_mainform.tpage_setup.get_output_folder_from_text_box();
                Process.Start(@"" + Output_folder);
            }
            catch (System.Exception)
            {

                MessageBox.Show("Please specify the output folder in the project settings page!");
            }
        }

        public void Fill_combobox_segments()
        {
            comboBox_segment_name.Items.Clear();
            if (_AGEN_mainform.lista_segments != null && _AGEN_mainform.lista_segments.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _AGEN_mainform.lista_segments.Count; ++i)
                    {
                        comboBox_segment_name.Items.Add(_AGEN_mainform.lista_segments[i]);
                    }
                    comboBox_segment_name.SelectedIndex = 0;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public void set_combobox_segment_name()
        {
            comboBox_segment_name.SelectedIndex = comboBox_segment_name.Items.IndexOf(_AGEN_mainform.current_segment);
        }

        private void ComboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.current_segment = comboBox_segment_name.Text;
            _AGEN_mainform.tpage_setup.set_combobox_segment_name();


        }
    }
}
