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
    public partial class AGEN_Viewport_Settings : Form
    {


        private ContextMenuStrip ContextMenuStrip_bands;
        private ContextMenuStrip ContextMenuStrip_blocks;
        bool is_main_view_picked = false;
        string col_bn = "Block Name";
        string col_rot = "Rotation";
        string col_space = "Location";
        string col_x = "X Paper Space";
        string col_y = "Y Paper Space";
        string col_pos = "Block Position";
        bool Template_is_open = false;


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_add_to_list);
            lista_butoane.Add(button_align_config_saveall);
            lista_butoane.Add(button_browser_dwt);
            lista_butoane.Add(button_browse_north_arrow);
            lista_butoane.Add(button_close_template2);
            lista_butoane.Add(button_draw_bands);
            lista_butoane.Add(button_remove_band);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_add_to_list);
            lista_butoane.Add(button_align_config_saveall);
            lista_butoane.Add(button_browser_dwt);
            lista_butoane.Add(button_browse_north_arrow);
            lista_butoane.Add(button_close_template2);
            lista_butoane.Add(button_draw_bands);
            lista_butoane.Add(button_remove_band);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public AGEN_Viewport_Settings()
        {
            InitializeComponent();
            _AGEN_mainform.nume_main_vp = comboBox_bands.Items[1].ToString();
            _AGEN_mainform.nume_banda_prof = comboBox_bands.Items[2].ToString();
            _AGEN_mainform.nume_banda_prop = comboBox_bands.Items[3].ToString();
            _AGEN_mainform.nume_banda_cross = comboBox_bands.Items[4].ToString();
            _AGEN_mainform.nume_banda_mat = comboBox_bands.Items[5].ToString();
            _AGEN_mainform.nume_banda_profband = comboBox_bands.Items[7].ToString();
            _AGEN_mainform.nume_banda_tblk_band = comboBox_bands.Items[8].ToString();

            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Pick band" };
            toolStripMenuItem1.Click += button_define_one_band_Click;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Remove band" };
            toolStripMenuItem2.Click += button_remove_band_Click;

            ContextMenuStrip_bands = new ContextMenuStrip();
            ContextMenuStrip_bands.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2 });

            var toolStripMenuItem3 = new ToolStripMenuItem { Text = "Pick block" };
            toolStripMenuItem3.Click += button_define_one_block_Click;

            var toolStripMenuItem4 = new ToolStripMenuItem { Text = "Remove block" };
            toolStripMenuItem4.Click += button_remove_one_block_Click;

            ContextMenuStrip_blocks = new ContextMenuStrip();
            ContextMenuStrip_blocks.Items.AddRange(new ToolStripItem[] { toolStripMenuItem3, toolStripMenuItem4 });

        }


        private void dataGridView_bands_Click(object sender, EventArgs e)
        {
            Type t = e.GetType();
            if (t.Equals(typeof(MouseEventArgs)))
            {
                MouseEventArgs mouse = (MouseEventArgs)e;
                if (mouse.Button == MouseButtons.Right)
                {



                    ContextMenuStrip_bands.Show(Cursor.Position);
                    ContextMenuStrip_bands.Visible = true;




                }
            }
            else
            {
                ContextMenuStrip_bands.Visible = false;
            }
        }

        private void dataGridView_bands_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_bands.CurrentCell = dataGridView_bands.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_bands.Show(Cursor.Position);
                ContextMenuStrip_bands.Visible = true;
            }
            else
            {
                ContextMenuStrip_bands.Visible = false;
            }
        }


        private void dataGridView_blocks_Click(object sender, EventArgs e)
        {
            Type t = e.GetType();
            if (t.Equals(typeof(MouseEventArgs)))
            {
                MouseEventArgs mouse = (MouseEventArgs)e;
                if (mouse.Button == MouseButtons.Right)
                {



                    ContextMenuStrip_blocks.Show(Cursor.Position);
                    ContextMenuStrip_blocks.Visible = true;




                }
            }
            else
            {
                ContextMenuStrip_blocks.Visible = false;
            }
        }

        private void dataGridView_blocks_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_bands.CurrentCell = dataGridView_bands.Rows[e.RowIndex].Cells[e.ColumnIndex];
                ContextMenuStrip_blocks.Show(Cursor.Position);
                ContextMenuStrip_blocks.Visible = true;
            }
            else
            {
                ContextMenuStrip_blocks.Visible = false;
            }
        }


        public string get_template_name_from_text_box()
        {
            return textBox_template_name.Text;
        }

        public string get_prefix_name_from_text_box()
        {
            return textBox_prefix_name.Text;
        }

        public string get_suffix_name_from_text_box()
        {
            return textBox_suffix.Text;
        }

        public string get_start_number_from_text_box()
        {
            return textBox_name_start_number.Text;
        }

        public string get_increment_from_text_box()
        {
            return textBox_name_increment.Text;
        }

        public string Get_combobox_viewport_scale_text()
        {
            return comboBox_vw_scale.Text;
        }
        public void Set_combobox_units_to_m()
        {
            comboBox_dwgunits.SelectedIndex = 1;
        }
        public void Set_combobox_units_to_ft()
        {
            comboBox_dwgunits.SelectedIndex = 0;
        }

        public string Get_combobox_units_text()
        {
            return comboBox_dwgunits.Text;
        }

        public void set_textBox_template_name(string dwt_name)
        {
            textBox_template_name.Text = dwt_name;
        }

        public void Set_prefix_text_box(string Prefix)
        {
            textBox_prefix_name.Text = Prefix;
        }

        public void Set_suffix_text_box(string Suffix)
        {
            textBox_suffix.Text = Suffix;
        }

        public void Set_start_no_text_box(string Startno)
        {
            textBox_name_start_number.Text = Startno;
        }

        public void Set_increment_text_box(string Increment)
        {
            textBox_name_increment.Text = Increment;
        }


        public int Get_combobox_viewport_scale_count()
        {
            return comboBox_vw_scale.Items.Count;
        }

        public string Get_combobox_viewport_scale(int sel_index)
        {
            return comboBox_vw_scale.Items[sel_index].ToString();
        }

        public void Set_combobox_viewport_scale(int sel_index)
        {
            comboBox_vw_scale.SelectedIndex = sel_index;
        }

        public void Set_content_of_combobox_viewport_scale()
        {
            comboBox_vw_scale.Items.Clear();
            if (_AGEN_mainform.units_of_measurement == "f")
            {
                string inch = "\u0022";
                comboBox_vw_scale.Items.Add("1");
                comboBox_vw_scale.Items.Add("1" + inch + "=10'");
                comboBox_vw_scale.Items.Add("1" + inch + "=20'");
                comboBox_vw_scale.Items.Add("1" + inch + "=30'");
                comboBox_vw_scale.Items.Add("1" + inch + "=40'");
                comboBox_vw_scale.Items.Add("1" + inch + "=50'");
                comboBox_vw_scale.Items.Add("1" + inch + "=60'");
                comboBox_vw_scale.Items.Add("1" + inch + "=100'");
                comboBox_vw_scale.Items.Add("1" + inch + "=200'");
                comboBox_vw_scale.Items.Add("1" + inch + "=300'");
                comboBox_vw_scale.Items.Add("1" + inch + "=400'");
                comboBox_vw_scale.Items.Add("1" + inch + "=500'");
                comboBox_vw_scale.Items.Add("1" + inch + "=600'");
                comboBox_vw_scale.Items.Add("1" + inch + "=700'");
                comboBox_vw_scale.Items.Add("1" + inch + "=800'");
                comboBox_vw_scale.Items.Add("1" + inch + "=900'");
                comboBox_vw_scale.Items.Add("1" + inch + "=1000'");

            }
            else
            {
                comboBox_vw_scale.Items.Add("1:500");
                comboBox_vw_scale.Items.Add("1:750");
                comboBox_vw_scale.Items.Add("1:1000");
                comboBox_vw_scale.Items.Add("1:2000");
                comboBox_vw_scale.Items.Add("1:2500");
                comboBox_vw_scale.Items.Add("1:5000");
                comboBox_vw_scale.Items.Add("1:7500");
                comboBox_vw_scale.Items.Add("1:10000");
            }
        }





        public string get_comboBox_viewport_target_areas(int index)
        {
            return comboBox_bands.Items[index].ToString();
        }



        private void button_close_template_Click(object sender, EventArgs e)
        {
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {


                set_enable_false();
                _AGEN_mainform.tpage_processing.Show();
                // Ag.WindowState = FormWindowState.Minimized;
                try
                {



                    string strTemplatePath = get_template_name_from_text_box();

                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;


                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {

                            Doc.CloseAndDiscard();



                        }

                    }
                    if (DocumentManager1.Count == 0)
                    {
                        string Template1 = "acad.dwt";
                        DocumentManager1.Add(Template1);
                    }



                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                set_enable_true();


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


                Ag.WindowState = FormWindowState.Normal;






            }
        }

        private void button_pick_all_blocks_location_PS_Click(object sender, EventArgs e)
        {
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            if (Ag != null && _AGEN_mainform.dt_blocks != null && _AGEN_mainform.dt_blocks.Rows.Count > 0)
            {



                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }

                if (comboBox_block_space.Text == "")
                {
                    MessageBox.Show("no modelspace or paper space specified");
                    set_enable_true();
                    return;
                }

                string strTemplatePath = get_template_name_from_text_box();
                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }



                bool Found1 = false;

                set_enable_false();
                try
                {





                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {

                            ThisDrawing = Doc;
                            DocumentManager1.MdiActiveDocument = ThisDrawing;
                            Found1 = true;
                        }
                    }

                    if (Found1 == false)
                    {


                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }


                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Ag.WindowState = FormWindowState.Minimized;

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                       
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                            for (int i = 0; i < _AGEN_mainform.dt_blocks.Rows.Count; i++)
                            {
                                if (_AGEN_mainform.dt_blocks.Rows[i][col_bn] != DBNull.Value &&
                                    _AGEN_mainform.dt_blocks.Rows[i][col_pos] != DBNull.Value &&
                                    Convert.ToString(_AGEN_mainform.dt_blocks.Rows[i][col_pos]) == "User Defined")
                                {
                                    string block_name = Convert.ToString(_AGEN_mainform.dt_blocks.Rows[i][col_bn]);
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the " + block_name + " insertion point");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Ag.WindowState = FormWindowState.Normal;
                                        return;
                                    }
                                    _AGEN_mainform.dt_blocks.Rows[i][col_x] = Point_res1.Value.X;
                                    _AGEN_mainform.dt_blocks.Rows[i][col_y] = Point_res1.Value.Y;

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
                set_enable_true();

                //_AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(false);


                Ag.WindowState = FormWindowState.Normal;
            }
        }

        private void button_define_one_block_Click(object sender, EventArgs e)
        {
            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

            if (Ag != null && _AGEN_mainform.dt_blocks != null && _AGEN_mainform.dt_blocks.Rows.Count > 0)
            {



                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }


                string strTemplatePath = get_template_name_from_text_box();
                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                Document ThisDrawing2 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    set_enable_true();
                    return;
                }



                bool Found1 = false;

                set_enable_false();

                using (DocumentLock lock2 = ThisDrawing2.LockDocument())
                {

                try
                {

                    int selected_index = -1;
                    string block_name = "";

                    int selected_index_data_grid = dataGridView_blocks.CurrentCell.RowIndex;
                    string data_grid_selected_block = Convert.ToString(dataGridView_blocks[0, selected_index_data_grid].Value);


                    for (int i = 0; i < _AGEN_mainform.dt_blocks.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.dt_blocks.Rows[i][col_bn] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.dt_blocks.Rows[i][col_bn]);
                            if (bn == data_grid_selected_block)
                            {
                                selected_index = i;
                                block_name = bn;
                            }
                        }
                    }

                    if (selected_index == -1)
                    {
                        set_enable_true();
                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                        Ag.WindowState = FormWindowState.Normal;
                        return;
                    }

                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {

                            ThisDrawing = Doc;
                            DocumentManager1.MdiActiveDocument = Doc;
                                
                            Found1 = true;
                            break;
                        }
                    }

                    if (Found1 == false)
                    {
                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                        DocumentManager1.MdiActiveDocument = ThisDrawing;
                    }


                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Ag.WindowState = FormWindowState.Minimized;

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {


                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);


                            if (Convert.ToString(_AGEN_mainform.dt_blocks.Rows[selected_index][col_pos]) == "User Defined")
                            {

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the " + block_name + " insertion point");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);


                                if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    Ag.WindowState = FormWindowState.Normal;
                                    return;
                                }
                                _AGEN_mainform.dt_blocks.Rows[selected_index][col_x] = Point_res1.Value.X;
                                _AGEN_mainform.dt_blocks.Rows[selected_index][col_y] = Point_res1.Value.Y;

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
                set_enable_true();

               


                Ag.WindowState = FormWindowState.Normal;
            }
        }

        private void button_remove_one_block_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {
                int selected_index = -1;
                string block_name = "";

                int selected_index_data_grid = dataGridView_blocks.CurrentCell.RowIndex;
                string data_grid_selected_block = Convert.ToString(dataGridView_blocks[0, selected_index_data_grid].Value);


                for (int i = 0; i < _AGEN_mainform.dt_blocks.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.dt_blocks.Rows[i][col_bn] != DBNull.Value)
                    {
                        string bn = Convert.ToString(_AGEN_mainform.dt_blocks.Rows[i][col_bn]);
                        if (bn == data_grid_selected_block)
                        {
                            selected_index = i;
                            block_name = bn;
                        }
                    }
                }

                if (selected_index == -1)
                {
                    set_enable_true();

                    return;
                }


                _AGEN_mainform.dt_blocks.Rows[selected_index].Delete();





              


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();

        }


        private void button_browser_dwt_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Template Files (*.dwt)|*.dwt|Drawing file (*.dwg)|*.dwg";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_template_name.Text = fbd.FileName;
                    _AGEN_mainform.template1 = fbd.FileName;
                }
            }
        }

        private void comboBox_dwgunits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[0].ToString())
            {
                _AGEN_mainform.units_of_measurement = "f";
            }
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[1].ToString())
            {
                _AGEN_mainform.units_of_measurement = "m";
            }
            Set_content_of_combobox_viewport_scale();
            _AGEN_mainform.tpage_setup.set_display_to_feet_or_meters();



        }

        private void comboBox_units_precision_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_units_precision.Text == "0")
            {
                _AGEN_mainform.round1 = 0;
            }
            else if (comboBox_units_precision.Text == "0.0")
            {
                _AGEN_mainform.round1 = 1;
            }
            else if (comboBox_units_precision.Text == "0.00")
            {
                _AGEN_mainform.round1 = 2;
            }
            else if (comboBox_units_precision.Text == "0.000")
            {
                _AGEN_mainform.round1 = 3;
            }
            else
            {
                _AGEN_mainform.round1 = 0;
            }
        }

        private void see_if_main_vp_is_selected()
        {
            is_main_view_picked = false;

            if (_AGEN_mainform.Vw_ps_x != 0 && _AGEN_mainform.Vw_ps_y != 0 && _AGEN_mainform.Vw_width > 0 && _AGEN_mainform.Vw_height > 0)
            {
                is_main_view_picked = true;

            }

        }

        public void creeaza_display_data_table(List<string> Lista1, List<string> Lista2, List<string> lista_string)
        {
            _AGEN_mainform.Data_Table_display_bands = new System.Data.DataTable();

            _AGEN_mainform.Data_Table_display_bands.Columns.Add("Band Name", typeof(string));
            _AGEN_mainform.Data_Table_display_bands.Columns.Add("Location Selected", typeof(string));

            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
            {
                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count == Lista1.Count)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);

                            _AGEN_mainform.Data_Table_display_bands.Rows.Add();
                            _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                            _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Location Selected"] = Lista1[i];
                        }
                    }
                }
            }

            if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count == Lista2.Count)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                        {

                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);

                            _AGEN_mainform.Data_Table_display_bands.Rows.Add();
                            _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                            _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Location Selected"] = Lista2[i];
                        }
                    }
                }
            }

            if (lista_string != null && _AGEN_mainform.Data_Table_extra_mainVP != null && lista_string.Count > 0)

            {
                for (int i = 0; i < _AGEN_mainform.Data_Table_extra_mainVP.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.Data_Table_extra_mainVP.Rows[i][0] != DBNull.Value)
                    {

                        string bn = Convert.ToString(_AGEN_mainform.Data_Table_extra_mainVP.Rows[i][0]);

                        _AGEN_mainform.Data_Table_display_bands.Rows.Add();
                        _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                        _AGEN_mainform.Data_Table_display_bands.Rows[_AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Location Selected"] = lista_string[i];
                    }
                }
            }

            dataGridView_bands.DataSource = _AGEN_mainform.Data_Table_display_bands;
            dataGridView_bands.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_bands.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_bands.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_bands.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_bands.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_bands.EnableHeadersVisualStyles = false;



        }



        private void button_align_config_saveall_Click(object sender, EventArgs e)
        {


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            if (Ag != null)
            {


                if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                {
                    _AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(true);

                }

            }



        }

        private void TextBox_keypress_only_pozitive_integers(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_integer_pozitive_at_keypress(sender, e);
        }

        private void button_add_to_list_Click(object sender, EventArgs e)
        {

            if ((comboBox_bands.SelectedIndex > 0 && comboBox_bands.SelectedIndex < 6) || comboBox_bands.SelectedIndex == 11)
            {
                bool exista = false;

                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (comboBox_bands.Text == CT)
                        {
                            exista = true;
                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        }
                    }
                }
                if (exista == false)
                {
                    _AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                    _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = comboBox_bands.Text;
                }

            }

            if (comboBox_bands.SelectedIndex == 6)
            {


                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {

                    foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                    {
                        if (Forma1 is Alignment_mdi.AGEN_custom_band_form)
                        {
                            Forma1.Focus();
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                            Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                              (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                            return;
                        }
                    }


                    try
                    {
                        Alignment_mdi.AGEN_custom_band_form forma2 = new Alignment_mdi.AGEN_custom_band_form();
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                        forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                             (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                    }
                    catch (System.Exception EX)
                    {
                        MessageBox.Show(EX.Message);
                    }

                }
                else
                {
                    MessageBox.Show("The main viewport is not picked");
                }
            }

            if (comboBox_bands.SelectedIndex == 7) //Profile as Band
            {
                bool exista = false;

                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (comboBox_bands.Text == CT)
                        {
                            exista = true;
                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        }
                    }
                }
                if (exista == false)
                {
                    _AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                    _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = comboBox_bands.Text;
                }
            }

            if (comboBox_bands.SelectedIndex == 8) //TBLK Band
            {
                bool exista = false;

                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (comboBox_bands.Text == CT)
                        {
                            exista = true;
                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        }
                    }
                }
                if (exista == false)
                {
                    _AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                    _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = comboBox_bands.Text;
                }
            }

            if (comboBox_bands.SelectedIndex == 9)//extra main VP
            {


                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {

                    foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                    {
                        if (Forma1 is Alignment_mdi.AGEN_extra_vp_form)
                        {
                            Forma1.Focus();
                            Forma1.WindowState = System.Windows.Forms.FormWindowState.Normal;
                            Forma1.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - Forma1.Width) / 2,
                              (Screen.PrimaryScreen.WorkingArea.Height - Forma1.Height) / 2);
                            return;
                        }
                    }


                    try
                    {
                        Alignment_mdi.AGEN_extra_vp_form forma2 = new Alignment_mdi.AGEN_extra_vp_form();
                        Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                        forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                             (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
                    }
                    catch (System.Exception EX)
                    {
                        MessageBox.Show(EX.Message);
                    }

                }
                else
                {
                    MessageBox.Show("The main viewport is not picked");
                }
            }



            if (comboBox_bands.SelectedIndex == 10) // 
            {

                bool exista = false;

                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (comboBox_bands.Text == CT)
                        {
                            exista = true;
                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        }
                    }
                }
                if (exista == false)
                {
                    _AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                    _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = comboBox_bands.Text;
                }
            }

            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
            {
                _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());
            }

        }

        public string get_combobox_units_precision()
        {
            return comboBox_units_precision.Text;
        }

        public void set_combobox_units_precision(string val)
        {
            if (comboBox_units_precision.Items.Contains(val) == true)
            {
                comboBox_units_precision.SelectedIndex = comboBox_units_precision.Items.IndexOf(val);
            }
            else
            {
                comboBox_units_precision.SelectedIndex = 0;
            }
        }

        private void button_remove_band_Click(object sender, EventArgs e)
        {


            set_enable_false();
            try
            {
                int selected_index = -1;
                bool is_regular_band = false;
                bool is_custom_band = false;

                string band_name = "";

                int selected_index_data_grid = dataGridView_bands.CurrentCell.RowIndex;

                string data_grid_selected_band = dataGridView_bands[0, selected_index_data_grid].Value.ToString();

                for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                    {
                        string bn = _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString();
                        if (bn == data_grid_selected_band)
                        {
                            selected_index = i;
                            is_custom_band = true;
                            band_name = bn;

                        }
                    }
                }

                for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                    {
                        string bn = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (bn == data_grid_selected_band)
                        {
                            selected_index = i;
                            is_regular_band = true;
                            band_name = bn;
                        }
                    }
                }

                bool proceseaza = false;

                if (is_regular_band == true || is_custom_band == true)
                {
                    proceseaza = true;
                }

                if (proceseaza == true)
                {
                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                    {
                        _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;



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

                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                        Microsoft.Office.Interop.Excel.Worksheet W_reg = null;
                        Microsoft.Office.Interop.Excel.Worksheet W_cust = null;

                        if (Workbook1.Worksheets.Count > 1)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet w2 in Workbook1.Worksheets)
                            {
                                if (w2.Name == "Regular_band_data")
                                {
                                    W_reg = w2;
                                }

                                if (w2.Name == "Custom_band_data")
                                {
                                    W_cust = w2;
                                }
                            }
                        }

                        try
                        {
                            bool one_reg_band_deleted = false;
                            bool one_cust_band_deleted = false;


                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {
                                if (band_name == _AGEN_mainform.nume_main_vp)
                                {

                                    //mainVP

                                    W1.Range["B36"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_main = false;
                                    is_main_view_picked = false;

                                    //crossing

                                    W1.Range["B37"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_cross = false;

                                    //ownership

                                    W1.Range["B38"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_owner = false;

                                    //profile

                                    W1.Range["B39"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_prof = false;

                                    //material
                                    W1.Range["B41"].Value = "False";

                                    _AGEN_mainform.Exista_viewport_mat = false;

                                    //profile band
                                    W1.Range["B42"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_prof_band = false;

                                    if (W_reg != null) W_reg.Delete();
                                    if (W_cust != null) W_cust.Delete();


                                    _AGEN_mainform.Data_Table_regular_bands = Functions.creeaza_regular_band_data_table_structure();
                                    _AGEN_mainform.Data_Table_custom_bands = Functions.creeaza_custom_band_data_table_structure();

                                }



                                else if (band_name == _AGEN_mainform.nume_banda_prof)
                                {
                                    //profile
                                    W1.Range["B39"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_prof = false;
                                    one_reg_band_deleted = true;
                                }
                                else if (band_name == _AGEN_mainform.nume_banda_profband)
                                {
                                    //profile band
                                    W1.Range["B42"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_prof_band = false;
                                    one_reg_band_deleted = true;
                                }

                                else if (band_name == _AGEN_mainform.nume_banda_prop)
                                {
                                    //owner
                                    W1.Range["B38"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_owner = false;
                                    one_reg_band_deleted = true;
                                }
                                else if (band_name == _AGEN_mainform.nume_banda_cross)
                                {
                                    //crossing
                                    W1.Range["B37"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_cross = false;
                                    one_reg_band_deleted = true;
                                }
                                else if (band_name == _AGEN_mainform.nume_banda_mat)
                                {
                                    //material
                                    W1.Range["B41"].Value = "False";
                                    _AGEN_mainform.Exista_viewport_mat = false;
                                    one_reg_band_deleted = true;
                                }

                                if (is_regular_band == true)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                    {
                                        if (selected_index >= 0)
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[selected_index].Delete();
                                            one_reg_band_deleted = true;
                                        }
                                    }
                                }


                                if (is_custom_band == true)
                                {
                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                                    {
                                        if (selected_index >= 0)
                                        {
                                            _AGEN_mainform.Data_Table_custom_bands.Rows[selected_index].Delete();
                                            one_cust_band_deleted = true;
                                        }
                                    }
                                }

                            }



                            if (one_reg_band_deleted == true) _AGEN_mainform.tpage_setup.transfera_regular_band_to_excel(Workbook1);
                            if (one_cust_band_deleted == true) _AGEN_mainform.tpage_setup.transfera_custom_band_to_excel(Workbook1);


                            Workbook1.Save();
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
                            if (W_reg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_reg);
                            if (W_cust != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cust);
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                        }
                        _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());
                    }
                }


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();

        }

        private void button_define_one_band_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {
                int selected_index = -1;
                bool is_regular_band = false;
                bool is_custom_band = false;
                bool is_multiple_prof = false;
                bool is_nodata_vp = false;
                int index_nodata = -1;

                string band_name = "";

                int selected_index_data_grid = dataGridView_bands.CurrentCell.RowIndex;

                string data_grid_selected_band = dataGridView_bands[0, selected_index_data_grid].Value.ToString();

                for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                    {
                        string bn = _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString();
                        if (bn == data_grid_selected_band)
                        {
                            selected_index = i;
                            is_custom_band = true;
                            band_name = bn;
                        }
                    }
                }

                for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                {
                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                    {
                        string bn = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                        if (bn == data_grid_selected_band)
                        {
                            selected_index = i;
                            is_regular_band = true;
                            band_name = bn;
                            if (band_name == Convert.ToString(comboBox_bands.Items[10]))
                            {
                                is_multiple_prof = true;
                            }
                            if (band_name == Convert.ToString(comboBox_bands.Items[11]))
                            {
                                is_nodata_vp = true;
                                index_nodata = i;
                            }
                        }
                    }
                }

                bool proceseaza = false;

                if ((is_regular_band == true && is_custom_band == false) || (is_regular_band == false && is_custom_band == true))
                {
                    proceseaza = true;
                }

                if (proceseaza == true)
                {
                    if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                    {
                        _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;



                        see_if_main_vp_is_selected();

                        string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();
                        DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                        if (System.IO.File.Exists(strTemplatePath) == false)
                        {
                            MessageBox.Show("template file not found");
                            set_enable_true();
                            return;
                        }

                        Template_is_open = false;
                        foreach (Document Doc in DocumentManager1)
                        {
                            if (Doc.Name == strTemplatePath)
                            {
                                Template_is_open = true;
                                ThisDrawing = Doc;
                                DocumentManager1.MdiActiveDocument = ThisDrawing;
                                Functions.Incarca_existing_Blocks_to_combobox(comboBox_existing_blocks);


                            }

                        }

                        if (Template_is_open == false)
                        {
                            ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                            Functions.Incarca_existing_Blocks_to_combobox(comboBox_existing_blocks);
                            Template_is_open = true;
                        }

                        string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();

                        if (Scale1.Contains(":") == true)
                        {
                            Scale1 = Scale1.Replace("1:", "");
                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
                            }
                        }

                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                        Ag.WindowState = FormWindowState.Minimized;

                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                double x1 = 0;
                                double y1 = 0;
                                double x2 = 0;
                                double y2 = 0;

                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                #region main viewport
                                if (band_name == _AGEN_mainform.nume_main_vp)
                                {




                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the plan view");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Ag.WindowState = FormWindowState.Normal;
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;

                                    Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                    Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\npick top right corner of the plan view");

                                    if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Ag.WindowState = FormWindowState.Normal;
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    x1 = Point_res1.Value.X;
                                    y1 = Point_res1.Value.Y;
                                    x2 = Point_res2.Value.X;
                                    y2 = Point_res2.Value.Y;

                                    if (y2 < y1)
                                    {
                                        double t1 = y1;
                                        y1 = y2;
                                        y2 = t1;

                                        if (x2 < x1)
                                        {
                                            double t2 = x1;
                                            x1 = x2;
                                            x2 = t2;
                                        }

                                    }

                                    _AGEN_mainform.Band_Separation = Math.Ceiling(3 * (Math.Abs(y2 - y1)) / 10) * 10;

                                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                    {
                                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                        if (_AGEN_mainform.nume_main_vp == CT)
                                        {
                                            _AGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                            _AGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                            _AGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                            _AGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_main_vp;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2 - x1);
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1 + x2) / 2;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = _AGEN_mainform.Band_Separation;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                            is_main_view_picked = true;
                                            _AGEN_mainform.Exista_viewport_main = true;


                                            i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                        }
                                    }
                                }
                                #endregion



                                if (is_nodata_vp == true)
                                {
                                    #region float band

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_float;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_float;
                                    PP_float = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the floating band");
                                    PP_float.AllowNone = false;
                                    Point_res_float = Editor1.GetPoint(PP_float);


                                    if (Point_res_float.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Ag.WindowState = FormWindowState.Normal;
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2_float;

                                    Alignment_mdi.Jig_rectangle_viewport_pick_points Jig_float = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                    Point_res2_float = Jig_float.StartJig(Point_res_float.Value, 1, "\npick top right corner of the floating band");

                                    if (Point_res2_float.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Ag.WindowState = FormWindowState.Normal;
                                        set_enable_true();
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }
                                    double x1_float = 0;
                                    double y1_float = 0;
                                    double x2_float = 0;
                                    double y2_float = 0;
                                    x1_float = Point_res_float.Value.X;
                                    y1_float = Point_res_float.Value.Y;
                                    x2_float = Point_res2_float.Value.X;
                                    y2_float = Point_res2_float.Value.Y;

                                    if (y2_float < y1_float)
                                    {
                                        double t1 = y1_float;
                                        y1_float = y2_float;
                                        y2_float = t1;

                                        if (x2_float < x1_float)
                                        {
                                            double t2 = x1_float;
                                            x1_float = x2_float;
                                            x2_float = t2;
                                        }

                                    }

                                    double h1 = Math.Abs(y2_float - y1_float);

                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["Custom_scale"] = 1;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["band_name"] = _AGEN_mainform.nume_banda_no_data;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_height"] = h1;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_width"] = Math.Abs(x2_float - x1_float);
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_ps_x"] = (x1_float + x2_float) / 2;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_ps_y"] = (y1_float + y2_float) / 2;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["band_separation"] = Math.Ceiling(3 * h1 / 10) * 10;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_ms_x"] = -1000;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_ms_y"] = -1000;
                                    _AGEN_mainform.Data_Table_regular_bands.Rows[index_nodata]["viewport_twist"] = 0;
                                    #endregion
                                }




                                #region multiple profile band vp
                                if (is_multiple_prof == true)
                                {
                                    int indexvp = 1;

                                    bool pick_vp = true;

                                    System.Data.DataTable dt1 = Functions.creeaza_regular_band_data_table_structure();
                                    bool pickpt = true;
                                    do
                                    {
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the profile band " + indexvp.ToString());
                                        PP1.AllowNone = false;
                                        Point_res1 = Editor1.GetPoint(PP1);

                                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            pick_vp = false;
                                            pickpt = false;
                                        }

                                        if (pickpt == true)
                                        {
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;

                                            Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\npick top right corner of the profile band " + indexvp.ToString());

                                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                pick_vp = false;
                                            }

                                            x1 = Point_res1.Value.X;
                                            y1 = Point_res1.Value.Y;
                                            x2 = Point_res2.Value.X;
                                            y2 = Point_res2.Value.Y;

                                            if (y2 < y1)
                                            {
                                                double t1 = y1;
                                                y1 = y2;
                                                y2 = t1;

                                                if (x2 < x1)
                                                {
                                                    double t2 = x1;
                                                    x1 = x2;
                                                    x2 = t2;
                                                }
                                            }


                                            dt1.Rows.Add();
                                            dt1.Rows[dt1.Rows.Count - 1]["viewport_height"] = Math.Abs(y2 - y1);
                                            dt1.Rows[dt1.Rows.Count - 1]["viewport_width"] = Math.Abs(x2 - x1);
                                            dt1.Rows[dt1.Rows.Count - 1]["viewport_ps_x"] = (x1 + x2) / 2;
                                            dt1.Rows[dt1.Rows.Count - 1]["viewport_ps_y"] = (y1 + y2) / 2;
                                        }



                                        ++indexvp;

                                    } while (pick_vp == true);

                                    for (int i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1; i >= 0; --i)
                                    {
                                        string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                        if (Convert.ToString(comboBox_bands.Items[10]) == CT)
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i].Delete();
                                        }
                                    }

                                    if (dt1.Rows.Count > 0)
                                    {
                                        for (int j = 0; j < dt1.Rows.Count; ++j)
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = Convert.ToString(comboBox_bands.Items[10]);
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_height"] = dt1.Rows[j]["viewport_height"];
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_width"] = dt1.Rows[j]["viewport_width"];
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_x"] = dt1.Rows[j]["viewport_ps_x"];
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_y"] = dt1.Rows[j]["viewport_ps_y"];
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_separation"] = DBNull.Value;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_x"] = DBNull.Value;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_y"] = DBNull.Value;
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_twist"] = DBNull.Value;

                                        }
                                    }


                                }

                                #endregion

                                if (is_main_view_picked == true)
                                {
                                    if (is_regular_band == true)
                                    {
                                        #region profile viewport
                                        if (band_name == _AGEN_mainform.nume_banda_prof)
                                        {

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of profile:");
                                            PP1.AllowNone = false;
                                            Point_res1 = Editor1.GetPoint(PP1);


                                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of profile:");

                                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }


                                            y1 = Point_res1.Value.Y;

                                            y2 = Point_res2.Value.Y;


                                            if (y2 < y1)
                                            {
                                                double t1 = y1;
                                                y1 = y2;
                                                y2 = t1;

                                            }

                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (_AGEN_mainform.nume_banda_prof == CT)
                                                {


                                                    _AGEN_mainform.Exista_viewport_prof = true;
                                                    _AGEN_mainform.Vw_prof_height = Math.Abs(y1 - y2);



                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_prof;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;
                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                }
                                            }
                                        }
                                        #endregion

                                        #region property viewport

                                        if (band_name == _AGEN_mainform.nume_banda_prop)
                                        {
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of ownership band:");
                                            PP1.AllowNone = false;
                                            Point_res1 = Editor1.GetPoint(PP1);

                                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of ownership band:");

                                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            y1 = Point_res1.Value.Y;
                                            y2 = Point_res2.Value.Y;

                                            if (y2 < y1)
                                            {
                                                double t1 = y1;
                                                y1 = y2;
                                                y2 = t1;
                                            }

                                            _AGEN_mainform.Point0_prop = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);

                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (band_name == CT)
                                                {


                                                    _AGEN_mainform.Exista_viewport_owner = true;
                                                    _AGEN_mainform.Vw_prop_height = Math.Abs(y1 - y2);

                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_prop;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_prop.X;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_prop.Y;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;


                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                }
                                            }
                                        }
                                        #endregion

                                        #region crossing viewport

                                        if (band_name == _AGEN_mainform.nume_banda_cross)
                                        {

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of crossing band:");
                                            PP1.AllowNone = false;
                                            Point_res1 = Editor1.GetPoint(PP1);


                                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of crossing band:");

                                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }


                                            y1 = Point_res1.Value.Y;
                                            y2 = Point_res2.Value.Y;

                                            _AGEN_mainform.Point0_cross = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);


                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (band_name == CT)
                                                {





                                                    _AGEN_mainform.Vw_cross_height = Math.Abs(y1 - y2);



                                                    _AGEN_mainform.Exista_viewport_cross = true;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_cross;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_cross.X;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_cross.Y;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;


                                                }
                                            }
                                        }
                                        #endregion

                                        #region material viewport

                                        if (band_name == _AGEN_mainform.nume_banda_mat)
                                        {

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of material band:");
                                            PP1.AllowNone = false;
                                            Point_res1 = Editor1.GetPoint(PP1);


                                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of material band:");

                                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            y1 = Point_res1.Value.Y;
                                            y2 = Point_res2.Value.Y;

                                            _AGEN_mainform.Point0_mat = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);

                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (band_name == CT)
                                                {

                                                    _AGEN_mainform.Vw_mat_height = Math.Abs(y1 - y2);


                                                    _AGEN_mainform.Exista_viewport_mat = true;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_mat;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_mat.X;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_mat.Y;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                }
                                            }
                                        }
                                        #endregion

                                        #region profile band viewport

                                        if (band_name == _AGEN_mainform.nume_banda_profband)
                                        {

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                                            PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of profile:");
                                            PP3.AllowNone = false;
                                            Point_res3 = Editor1.GetPoint(PP3);


                                            if (Point_res3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig4 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                            Point_res4 = Jig4.StartJig(Point_res3.Value, 1, "\nSpecify top of profile:");

                                            if (Point_res4.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            y1 = Point_res3.Value.Y;
                                            y2 = Point_res4.Value.Y;


                                            if (y2 < y1)
                                            {
                                                double t1 = y1;
                                                y1 = y2;
                                                y2 = t1;
                                            }

                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (_AGEN_mainform.nume_banda_profband == CT)
                                                {
                                                    _AGEN_mainform.Exista_viewport_prof_band = true;
                                                    _AGEN_mainform.Vw_profband_height = Math.Abs(y1 - y2);


                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_profband;

                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                }
                                            }








                                        }
                                        #endregion

                                        #region tblk band viewport


                                        if (band_name == _AGEN_mainform.nume_banda_tblk_band)
                                        {
                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                                            PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom left corner of tblk band:");
                                            PP5.AllowNone = false;
                                            Point_res5 = Editor1.GetPoint(PP5);

                                            if (Point_res5.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res6;
                                            Alignment_mdi.Jig_rectangle_viewport_pick_points Jig6 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                            Point_res6 = Jig6.StartJig(Point_res5.Value, 1, "\nSpecify top right corner of tblk band:");

                                            if (Point_res6.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                Ag.WindowState = FormWindowState.Normal;
                                                set_enable_true();
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }

                                            x1 = Point_res5.Value.X;
                                            x2 = Point_res6.Value.X;
                                            y1 = Point_res5.Value.Y;
                                            y2 = Point_res6.Value.Y;

                                            if (y2 < y1)
                                            {
                                                double t1 = y1;
                                                y1 = y2;
                                                y2 = t1;
                                            }



                                            Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_double = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify separation:");
                                            Prompt_double.AllowNegative = false;
                                            Prompt_double.AllowZero = true;
                                            Prompt_double.AllowNone = true;
                                            Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_dbl = ThisDrawing.Editor.GetDouble(Prompt_double);
                                            if (Rezultat_dbl.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                            {
                                                _AGEN_mainform.tblk_separation = Math.Abs(Rezultat_dbl.Value);
                                            }





                                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (band_name == CT)
                                                {
                                                    _AGEN_mainform.Exista_viewport_tblk = true;
                                                    _AGEN_mainform.Vw_tblk_height = Math.Abs(y1 - y2);


                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_tblk_band;

                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2 - x1);
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1 + x2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = _AGEN_mainform.tblk_separation;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = 0;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = 0;
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = 0;

                                                    i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;

                                                }
                                            }



                                        }

                                        #endregion


                                    }

                                    if (is_custom_band == true)
                                    {
                                        #region custom viewport

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of " + band_name + " band:");
                                        PP1.AllowNone = false;
                                        Point_res1 = Editor1.GetPoint(PP1);

                                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                        Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                        Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of " + band_name + " band:");

                                        if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        y1 = Point_res1.Value.Y;
                                        y2 = Point_res2.Value.Y;

                                        if (y2 < y1)
                                        {
                                            double t1 = y1;
                                            y1 = y2;
                                            y2 = t1;
                                        }


                                        for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                                        {
                                            string CT = _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"].ToString();
                                            if (band_name == CT)
                                            {
                                                Point3d Point0_c = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] = band_name;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"] = Point0_c.X;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"] = Point0_c.Y;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_twist"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"] = DBNull.Value;

                                                i = _AGEN_mainform.Data_Table_custom_bands.Rows.Count;
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                else if (is_multiple_prof == false && is_nodata_vp == false)
                                {
                                    MessageBox.Show("first you have to pick the main viewport!");
                                    set_enable_true();
                                    Ag.WindowState = FormWindowState.Normal;
                                    return;
                                }
                            }
                        }

                        #region region excel 
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

                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                        Microsoft.Office.Interop.Excel.Worksheet W_reg = null;
                        Microsoft.Office.Interop.Excel.Worksheet W_cust = null;

                        if (Workbook1.Worksheets.Count > 1)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet w2 in Workbook1.Worksheets)
                            {
                                if (w2.Name == "Regular_band_data")
                                {
                                    W_reg = w2;
                                }

                                if (w2.Name == "Custom_band_data")
                                {
                                    W_cust = w2;
                                }
                            }
                        }

                        try
                        {

                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {


                                if (is_regular_band == true)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                    {
                                        if (selected_index >= 0)
                                        {
                                            if (band_name == _AGEN_mainform.nume_main_vp)
                                            {
                                                //mainVP
                                                W1.Range["B10"].Value = _AGEN_mainform.Vw_ps_x;
                                                W1.Range["B11"].Value = _AGEN_mainform.Vw_ps_y;
                                                W1.Range["B12"].Value = _AGEN_mainform.Vw_width;
                                                W1.Range["B13"].Value = _AGEN_mainform.Vw_height;
                                                W1.Range["B36"].Value = "True";
                                            }
                                        }
                                    }
                                }
                            }



                            if (is_regular_band == true) _AGEN_mainform.tpage_setup.transfera_regular_band_to_excel(Workbook1);
                            if (is_custom_band == true) _AGEN_mainform.tpage_setup.transfera_custom_band_to_excel(Workbook1);


                            Workbook1.Save();
                            Workbook1.Close();
                            if (Excel1.Workbooks.Count == 0)
                            {
                                Excel1.Quit();
                            }
                            else
                            {
                                Excel1.Visible = true;
                            }
                            #endregion

                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                            if (W_reg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_reg);
                            if (W_cust != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cust);
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                        }
                        _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());
                        Ag.WindowState = FormWindowState.Normal;
                    }
                }


            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();





        }

        private void button_define_bands_Click(object sender, EventArgs e)
        {

            set_enable_false();
            try
            {

                string band_name = "";

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                {
                    _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;


                    see_if_main_vp_is_selected();

                    string strTemplatePath = _AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();
                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                    if (System.IO.File.Exists(strTemplatePath) == false)
                    {
                        MessageBox.Show("template file not found");
                        set_enable_true();
                        return;
                    }

                    Template_is_open = false;
                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {
                            Template_is_open = true;
                            ThisDrawing = Doc;
                            DocumentManager1.MdiActiveDocument = ThisDrawing;
                            Functions.Incarca_existing_Blocks_to_combobox(comboBox_existing_blocks);



                        }

                    }

                    if (Template_is_open == false)
                    {
                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                        Functions.Incarca_existing_Blocks_to_combobox(comboBox_existing_blocks);

                    }

                    string Scale1 = _AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();

                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _AGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                    }
                    else
                    {
                        if (Scale1.Contains(":") == true)
                        {
                            Scale1 = Scale1.Replace("1:", "");
                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _AGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
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
                                _AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                            }
                        }
                    }



                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    Ag.WindowState = FormWindowState.Minimized;

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            double x1 = 0;
                            double y1 = 0;
                            double x2 = 0;
                            double y2 = 0;

                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                            Functions.make_first_layout_active(Trans1, ThisDrawing.Database);


                            #region main viewport

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the plan view");
                            PP1.AllowNone = false;
                            Point_res1 = Editor1.GetPoint(PP1);


                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Ag.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;

                            Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\npick top right corner of the plan view");

                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Ag.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            x1 = Point_res1.Value.X;
                            y1 = Point_res1.Value.Y;
                            x2 = Point_res2.Value.X;
                            y2 = Point_res2.Value.Y;

                            if (y2 < y1)
                            {
                                double t1 = y1;
                                y1 = y2;
                                y2 = t1;

                                if (x2 < x1)
                                {
                                    double t2 = x1;
                                    x1 = x2;
                                    x2 = t2;
                                }

                            }

                            _AGEN_mainform.Band_Separation = Math.Ceiling(3 * (Math.Abs(y2 - y1)) / 10) * 10;


                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {

                                for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                {
                                    string CT = _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                    if (_AGEN_mainform.nume_main_vp == CT)
                                    {
                                        _AGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                        _AGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                        _AGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                        _AGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_main_vp;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2 - x1);
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1 + x2) / 2;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = _AGEN_mainform.Band_Separation;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                        is_main_view_picked = true;
                                        _AGEN_mainform.Exista_viewport_main = true;

                                        i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;


                                    }
                                }

                            }
                            else
                            {
                                _AGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                _AGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                _AGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                _AGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                _AGEN_mainform.Data_Table_regular_bands.Rows.Add();


                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = _AGEN_mainform.nume_main_vp;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_height"] = Math.Abs(y2 - y1);
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_width"] = Math.Abs(x2 - x1);
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_x"] = (x1 + x2) / 2;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_y"] = (y1 + y2) / 2;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_separation"] = _AGEN_mainform.Band_Separation;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_x"] = DBNull.Value;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_y"] = DBNull.Value;
                                _AGEN_mainform.Data_Table_regular_bands.Rows[_AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_twist"] = DBNull.Value;

                                is_main_view_picked = true;
                                _AGEN_mainform.Exista_viewport_main = true;
                            }


                            #endregion


                            bool is_no_data_band = false;

                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                            {
                                if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                {
                                    band_name = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                                    if (band_name == _AGEN_mainform.nume_banda_no_data)
                                    {
                                        #region no data band

                                        is_no_data_band = true;
                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res_float;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP_float;
                                        PP_float = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the no data band");
                                        PP_float.AllowNone = false;
                                        Point_res_float = Editor1.GetPoint(PP_float);


                                        if (Point_res_float.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2_float;

                                        Alignment_mdi.Jig_rectangle_viewport_pick_points Jig_float = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                        Point_res2_float = Jig_float.StartJig(Point_res_float.Value, 1, "\npick top right corner of the no data band");

                                        if (Point_res2_float.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }
                                        double x1_float = 0;
                                        double y1_float = 0;
                                        double x2_float = 0;
                                        double y2_float = 0;
                                        x1_float = Point_res_float.Value.X;
                                        y1_float = Point_res_float.Value.Y;
                                        x2_float = Point_res2_float.Value.X;
                                        y2_float = Point_res2_float.Value.Y;

                                        if (y2_float < y1_float)
                                        {
                                            double t1 = y1_float;
                                            y1_float = y2_float;
                                            y2_float = t1;

                                            if (x2_float < x1_float)
                                            {
                                                double t2 = x1_float;
                                                x1_float = x2_float;
                                                x2_float = t2;
                                            }

                                        }

                                        double h1 = Math.Abs(y2_float - y1_float);

                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_no_data;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = h1;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2_float - x1_float);
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1_float + x2_float) / 2;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1_float + y2_float) / 2;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = Math.Ceiling(3 * h1 / 10) * 10;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = -1000;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = -1000;
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = 0;

                                        i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;


                                        #endregion
                                    }
                                }
                            }



                            if (is_main_view_picked == true)
                            {

                                for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                    {
                                        band_name = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                                        if (band_name != _AGEN_mainform.nume_main_vp)
                                        {
                                            #region profile viewport
                                            if (band_name == _AGEN_mainform.nume_banda_prof)
                                            {

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                                                PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of profile:");
                                                PP3.AllowNone = false;
                                                Point_res3 = Editor1.GetPoint(PP3);


                                                if (Point_res3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig4 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                                Point_res4 = Jig4.StartJig(Point_res3.Value, 1, "\nSpecify top of profile:");

                                                if (Point_res4.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                y1 = Point_res3.Value.Y;
                                                y2 = Point_res4.Value.Y;


                                                if (y2 < y1)
                                                {
                                                    double t1 = y1;
                                                    y1 = y2;
                                                    y2 = t1;
                                                }


                                                _AGEN_mainform.Exista_viewport_prof = true;
                                                _AGEN_mainform.Vw_prof_height = Math.Abs(y1 - y2);



                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_prof;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;



                                            }
                                            #endregion

                                            #region property viewport

                                            if (band_name == _AGEN_mainform.nume_banda_prop)
                                            {
                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                                                PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of ownership band:");
                                                PP5.AllowNone = false;
                                                Point_res5 = Editor1.GetPoint(PP5);

                                                if (Point_res5.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res6;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig6 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                                Point_res6 = Jig6.StartJig(Point_res5.Value, 1, "\nSpecify top of ownership band:");

                                                if (Point_res6.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                y1 = Point_res5.Value.Y;
                                                y2 = Point_res6.Value.Y;

                                                if (y2 < y1)
                                                {
                                                    double t1 = y1;
                                                    y1 = y2;
                                                    y2 = t1;
                                                }

                                                _AGEN_mainform.Point0_prop = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);

                                                _AGEN_mainform.Exista_viewport_owner = true;
                                                _AGEN_mainform.Vw_prop_height = Math.Abs(y1 - y2);


                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_prop;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_prop.X;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_prop.Y;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                            }

                                            #endregion

                                            #region crossing viewport

                                            if (band_name == _AGEN_mainform.nume_banda_cross)
                                            {

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res7;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP7;
                                                PP7 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of crossing band:");
                                                PP7.AllowNone = false;
                                                Point_res7 = Editor1.GetPoint(PP7);


                                                if (Point_res7.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res8;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig8 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                                Point_res8 = Jig8.StartJig(Point_res7.Value, 1, "\nSpecify top of crossing band:");

                                                if (Point_res8.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }


                                                y1 = Point_res7.Value.Y;
                                                y2 = Point_res8.Value.Y;



                                                _AGEN_mainform.Point0_cross = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);
                                                _AGEN_mainform.Vw_cross_height = Math.Abs(y1 - y2);



                                                _AGEN_mainform.Exista_viewport_cross = true;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_cross;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_cross.X;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_cross.Y;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                            }
                                            #endregion

                                            #region material viewport

                                            if (band_name == _AGEN_mainform.nume_banda_mat)
                                            {

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res9;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP9;
                                                PP9 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of material band:");
                                                PP9.AllowNone = false;
                                                Point_res9 = Editor1.GetPoint(PP9);


                                                if (Point_res9.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res10;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig10 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                                Point_res10 = Jig10.StartJig(Point_res9.Value, 1, "\nSpecify top of material band:");

                                                if (Point_res10.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                y1 = Point_res9.Value.Y;
                                                y2 = Point_res10.Value.Y;

                                                _AGEN_mainform.Point0_mat = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);
                                                _AGEN_mainform.Vw_mat_height = Math.Abs(y1 - y2);

                                                _AGEN_mainform.Exista_viewport_mat = true;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_mat;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = _AGEN_mainform.Point0_mat.X;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = _AGEN_mainform.Point0_mat.Y;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                            }
                                            #endregion

                                            #region profile band viewport

                                            if (band_name == _AGEN_mainform.nume_banda_profband)
                                            {

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res3;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP3;
                                                PP3 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of profile band:");
                                                PP3.AllowNone = false;
                                                Point_res3 = Editor1.GetPoint(PP3);


                                                if (Point_res3.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res4;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig4 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                                Point_res4 = Jig4.StartJig(Point_res3.Value, 1, "\nSpecify top of profile band:");

                                                if (Point_res4.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                y1 = Point_res3.Value.Y;
                                                y2 = Point_res4.Value.Y;


                                                if (y2 < y1)
                                                {
                                                    double t1 = y1;
                                                    y1 = y2;
                                                    y2 = t1;
                                                }


                                                _AGEN_mainform.Exista_viewport_prof_band = true;
                                                _AGEN_mainform.Vw_profband_height = Math.Abs(y1 - y2);


                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _AGEN_mainform.Vw_scale;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_profband;

                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;


                                            }
                                            #endregion

                                            #region tblk band viewport


                                            if (band_name == _AGEN_mainform.nume_banda_tblk_band)
                                            {
                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res5;
                                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP5;
                                                PP5 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left corner of the tblk band:");
                                                PP5.AllowNone = false;
                                                Point_res5 = Editor1.GetPoint(PP5);

                                                if (Point_res5.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }

                                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res6;
                                                Alignment_mdi.Jig_rectangle_viewport_pick_points Jig6 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                                Point_res6 = Jig6.StartJig(Point_res5.Value, 1, "\nspecify the top right corner of the tblk band:");

                                                if (Point_res6.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    Ag.WindowState = FormWindowState.Normal;
                                                    set_enable_true();
                                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                    return;
                                                }
                                                x1 = Point_res5.Value.X;
                                                x2 = Point_res6.Value.X;
                                                y1 = Point_res5.Value.Y;
                                                y2 = Point_res6.Value.Y;

                                                if (y2 < y1)
                                                {
                                                    double t1 = y1;
                                                    y1 = y2;
                                                    y2 = t1;
                                                }



                                                Autodesk.AutoCAD.EditorInput.PromptDoubleOptions Prompt_double = new Autodesk.AutoCAD.EditorInput.PromptDoubleOptions("\n" + "Specify separation:");
                                                Prompt_double.AllowNegative = false;
                                                Prompt_double.AllowZero = true;
                                                Prompt_double.AllowNone = true;
                                                Autodesk.AutoCAD.EditorInput.PromptDoubleResult Rezultat_dbl = ThisDrawing.Editor.GetDouble(Prompt_double);
                                                if (Rezultat_dbl.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                                {
                                                    _AGEN_mainform.tblk_separation = Math.Abs(Rezultat_dbl.Value);
                                                }





                                                _AGEN_mainform.Exista_viewport_tblk = true;
                                                _AGEN_mainform.Vw_tblk_height = Math.Abs(y1 - y2);


                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _AGEN_mainform.nume_banda_tblk_band;

                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2 - x1);
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1 + x2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = _AGEN_mainform.tblk_separation;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = 0;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = 0;
                                                _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = 0;

                                            }

                                            #endregion
                                        }
                                    }
                                }


                                for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                                {
                                    if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                                    {
                                        band_name = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);

                                        #region custom viewport

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res11;
                                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP11;
                                        PP11 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of " + band_name + " band:");
                                        PP11.AllowNone = false;
                                        Point_res11 = Editor1.GetPoint(PP11);

                                        if (Point_res11.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res12;
                                        Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig12 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                        Point_res12 = Jig12.StartJig(Point_res11.Value, 1, "\nSpecify top of " + band_name + " band:");

                                        if (Point_res12.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                        {
                                            Ag.WindowState = FormWindowState.Normal;
                                            set_enable_true();
                                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                            return;
                                        }

                                        y1 = Point_res11.Value.Y;
                                        y2 = Point_res12.Value.Y;

                                        if (y2 < y1)
                                        {
                                            double t1 = y1;
                                            y1 = y2;
                                            y2 = t1;
                                        }
                                        Point3d Point0_c = new Point3d(0, Functions.calculate_vp_ms_y(_AGEN_mainform.Band_Separation, y2), 0);

                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] = band_name;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_width"] = _AGEN_mainform.Vw_width;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ps_x"] = _AGEN_mainform.Vw_ps_x;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_separation"] = DBNull.Value;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_x"] = Point0_c.X;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_ms_y"] = Point0_c.Y;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["viewport_twist"] = DBNull.Value;
                                        _AGEN_mainform.Data_Table_custom_bands.Rows[i]["Custom_scale"] = DBNull.Value;

                                        #endregion
                                    }
                                }
                            }
                            else if (is_no_data_band == false)
                            {
                                MessageBox.Show("please add the plan view first!");
                                Ag.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                        }
                    }



                    #region region excel 

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

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_AGEN_mainform.config_path);

                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                    Microsoft.Office.Interop.Excel.Worksheet W_reg = null;
                    Microsoft.Office.Interop.Excel.Worksheet W_cust = null;

                    if (Workbook1.Worksheets.Count > 1)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet w2 in Workbook1.Worksheets)
                        {
                            if (w2.Name == "Regular_band_data")
                            {
                                W_reg = w2;
                            }

                            if (w2.Name == "Custom_band_data")
                            {
                                W_cust = w2;
                            }
                        }
                    }

                    try
                    {

                        if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                        {
                            if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {
                                if (band_name == _AGEN_mainform.nume_main_vp)
                                {
                                    //mainVP
                                    W1.Range["B36"].Value = "True";
                                }
                            }
                        }



                        _AGEN_mainform.tpage_setup.transfera_regular_band_to_excel(Workbook1);
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0) _AGEN_mainform.tpage_setup.transfera_custom_band_to_excel(Workbook1);


                        Workbook1.Save();
                        Workbook1.Close();
                        if (Excel1.Workbooks.Count == 0)
                        {
                            Excel1.Quit();
                        }
                        else
                        {
                            Excel1.Visible = true;
                        }
                        #endregion

                    }
                    catch (System.Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                        if (W_reg != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_reg);
                        if (W_cust != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_cust);
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }
                    _AGEN_mainform.tpage_viewport_settings.creeaza_display_data_table(Functions.Creaza_lista_regular_vp_picked(), Functions.Creaza_lista_custom_vp_picked(), Functions.Creaza_lista_custom_vp_extra_picked());
                    Ag.WindowState = FormWindowState.Normal;
                }



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();






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

        private void label_sheet_naming_Click(object sender, EventArgs e)
        {
            if (Functions.is_dan_popescu() == true)
            {
                if (comboBox_segment_name.Visible == false)
                {
                    comboBox_segment_name.Visible = true;
                    label_segm_name.Visible = true;
                    panel_sheet_naming.Size = new Size(366, 141);
                }
                else
                {
                    comboBox_segment_name.Visible = false;
                    label_segm_name.Visible = false;
                    panel_sheet_naming.Size = new Size(366, 111);
                }
            }
        }

        private void radioButton_lr_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_lr.Checked == true)
            {
                _AGEN_mainform.Left_to_Right = true;
            }
            else
            {
                _AGEN_mainform.Left_to_Right = false;
            }
        }

        public void set_radioButton_left_right(bool bool1)
        {
            if (bool1 == false)
            {
                radioButton_rl.Checked = true;
            }
            else
            {
                radioButton_lr.Checked = true;
            }
        }

        public string get_comboBox_bands_multiple_vp()
        {
            return Convert.ToString(comboBox_bands.Items[10]);
        }



        private void button_add_block_Click(object sender, EventArgs e)
        {
            if (comboBox_existing_blocks.Text != "" && comboBox_block_space.Text != "")
            {


                if (_AGEN_mainform.dt_blocks == null)
                {
                    _AGEN_mainform.dt_blocks = new System.Data.DataTable();

                    _AGEN_mainform.dt_blocks.Columns.Add(col_bn, typeof(string));
                    _AGEN_mainform.dt_blocks.Columns.Add(col_rot, typeof(string));
                    _AGEN_mainform.dt_blocks.Columns.Add(col_pos, typeof(string));
                    _AGEN_mainform.dt_blocks.Columns.Add(col_x, typeof(double));
                    _AGEN_mainform.dt_blocks.Columns.Add(col_y, typeof(double));
                    _AGEN_mainform.dt_blocks.Columns.Add(col_space, typeof(string));
                }

                _AGEN_mainform.dt_blocks.Rows.Add();
                _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_bn] = comboBox_existing_blocks.Text;

                if (radioButton_rot_0.Checked == true)
                {
                    _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_rot] = "0";
                }
                else
                {
                    _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_rot] = "Sheet Index";
                }

                if (radioButton_block_user_defined.Checked == true)
                {
                    _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_pos] = "User Defined";
                }
                else
                {
                    _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_pos] = "At Matchlines";
                }

                _AGEN_mainform.dt_blocks.Rows[_AGEN_mainform.dt_blocks.Rows.Count - 1][col_space] = comboBox_block_space.Text;


            }

            display_dt_blocks();
        }

        public void display_dt_blocks()
        {
            dataGridView_blocks.DataSource = _AGEN_mainform.dt_blocks;
            dataGridView_blocks.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_blocks.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_blocks.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_blocks.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_blocks.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_blocks.EnableHeadersVisualStyles = false;
        }

        private void comboBox_existing_blocks_DropDown(object sender, EventArgs e)
        {
            Functions.Incarca_existing_Blocks_to_combobox(comboBox_existing_blocks);
        }
    }
}
