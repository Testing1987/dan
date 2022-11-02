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
        bool Freeze_operations = false;

        Point3d custom_band_zero_point = new Point3d(30000, -1000, 0);

        public string nume_banda_prof = "";
        public string nume_banda_prop = "";
        public string nume_banda_cross = "";
        public string nume_banda_mat = "";
        public string nume_main_vp = "";

        bool is_main_view_picked = false;

        public AGEN_Viewport_Settings()
        {
            InitializeComponent();
            nume_main_vp = comboBox_bands_target_areas.Items[1].ToString();
            nume_banda_prof = comboBox_bands_target_areas.Items[2].ToString();
            nume_banda_prop = comboBox_bands_target_areas.Items[3].ToString();
            nume_banda_cross = comboBox_bands_target_areas.Items[4].ToString();
            nume_banda_mat = comboBox_bands_target_areas.Items[5].ToString();
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

        public void set_dataGridView_north_arrow_blocks()
        {
            dataGridView_north_arrow_blocks.DataSource = AGEN_mainform.Data_table_blocks;
            AGEN_mainform.tpage_viewport_settings.dataGridView_north_arrow_blocks.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }



        public string get_comboBox_viewport_target_areas(int index)
        {
            return comboBox_bands_target_areas.Items[index].ToString();
        }



        private void button_close_template_Click(object sender, EventArgs e)
        {
            AGEN_mainform Ag = this.MdiParent as AGEN_mainform;
            if (Ag != null)
            {

                if (Freeze_operations == false)
                {
                    Freeze_operations = true;
                    AGEN_mainform.tpage_processing.Show();
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
                    Freeze_operations = false;
                    AGEN_mainform.Template_is_open = false;

                    AGEN_mainform.tpage_processing.Hide();
                    AGEN_mainform.tpage_blank.Hide();
                    AGEN_mainform.tpage_setup.Hide();
                  
                    AGEN_mainform.tpage_tblk_attrib.Hide();
                    AGEN_mainform.tpage_sheetindex.Hide();
                    AGEN_mainform.tpage_layer_alias.Hide();
                    AGEN_mainform.tpage_crossing_scan.Hide();
                    AGEN_mainform.tpage_crossing_draw.Hide();
                    AGEN_mainform.tpage_profilescan.Hide();
                    AGEN_mainform.tpage_profdraw.Hide();
                    AGEN_mainform.tpage_proflabel.Hide();
                    AGEN_mainform.tpage_owner_scan.Hide();
                    AGEN_mainform.tpage_owner_draw.Hide();
                    AGEN_mainform.tpage_mat.Hide();
                    AGEN_mainform.tpage_cust_scan.Hide();
                    AGEN_mainform.tpage_cust_draw.Hide();
                    AGEN_mainform.tpage_al_gen.Hide();
                    AGEN_mainform.tpage_asbuilt.Hide();

                    AGEN_mainform.tpage_viewport_settings.Show();
           

                    Ag.WindowState = FormWindowState.Normal;




                }

            }
        }

        private void button_browse_north_arrow_Click(object sender, EventArgs e)
        {
            AGEN_mainform Ag = this.MdiParent as AGEN_mainform;
            if (Ag != null)
            {

                AGEN_mainform.locatie_config_file = Ag.get_textBox_config_file_location();

                if (System.IO.File.Exists(AGEN_mainform.locatie_config_file) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }

                if (comboBox_type_of_block.Text == "")
                {
                    MessageBox.Show("no type of block specified");
                    Freeze_operations = false;
                    return;
                }

                string strTemplatePath = get_template_name_from_text_box();
                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                if (System.IO.File.Exists(strTemplatePath) == false)
                {
                    MessageBox.Show("template file not found");
                    Freeze_operations = false;
                    return;
                }



                bool Found1 = false;

                if (Freeze_operations == false)
                {
                    Freeze_operations = true;
                    try
                    {

                        if (dataGridView_north_arrow_blocks.Rows.Count == 0)
                        {
                            AGEN_mainform.Data_table_blocks = new System.Data.DataTable();
                            AGEN_mainform.Data_table_blocks.Columns.Add("TYPE", typeof(String));
                            AGEN_mainform.Data_table_blocks.Columns.Add("BLOCK_NAME", typeof(String));
                            AGEN_mainform.Data_table_blocks.Columns.Add("SCALE", typeof(double));
                            AGEN_mainform.Data_table_blocks.Columns.Add("X", typeof(double));
                            AGEN_mainform.Data_table_blocks.Columns.Add("Y", typeof(double));

                        }




                        foreach (Document Doc in DocumentManager1)
                        {
                            if (Doc.Name == strTemplatePath)
                            {
                                AGEN_mainform.Template_is_open = true;
                                ThisDrawing = Doc;
                                DocumentManager1.CurrentDocument = ThisDrawing;
                                Found1 = true;
                            }
                        }

                        if (Found1 == false)
                        {
                            AGEN_mainform.Template_is_open = false;

                            ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                            DocumentManager1.CurrentDocument = ThisDrawing;
                            Functions.Incarca_existing_Blocks_to_combobox(comboBox_blocks_NA);
                            if (comboBox_blocks_NA.Items.Count > 1) comboBox_blocks_NA.Items.Insert(1, AGEN_mainform.insertNAtoMS);
                            if (comboBox_blocks_NA.Items.Count == 0)
                            {
                                comboBox_blocks_NA.Items.Add("");
                                comboBox_blocks_NA.Items.Add(AGEN_mainform.insertNAtoMS);
                            }
                            MessageBox.Show("the template file has been open, please select your north arrow block first");
                            Freeze_operations = false;
                            AGEN_mainform.Template_is_open = true;

                            return;
                        }


                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        Ag.WindowState = FormWindowState.Minimized;

                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {

                                if (comboBox_type_of_block.Text == "North Arrow")
                                {

                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the north arrow insertion point");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Ag.WindowState = FormWindowState.Normal;
                                        return;
                                    }



                                    AGEN_mainform.NA_x = Point_res1.Value.X;
                                    AGEN_mainform.NA_y = Point_res1.Value.Y;
                                    AGEN_mainform.NA_name = comboBox_blocks_NA.Text;


                                    string Block_type = comboBox_type_of_block.Text;
                                    bool Exista = false;

                                    if (Block_type != "")
                                    {

                                        if (AGEN_mainform.Data_table_blocks.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_table_blocks.Rows.Count; ++i)
                                            {
                                                string BT = AGEN_mainform.Data_table_blocks.Rows[i][0].ToString();
                                                if (Block_type == BT)
                                                {

                                                    AGEN_mainform.Data_table_blocks.Rows[i]["TYPE"] = BT;
                                                    AGEN_mainform.Data_table_blocks.Rows[i][AGEN_mainform.Col_x] = AGEN_mainform.NA_x;
                                                    AGEN_mainform.Data_table_blocks.Rows[i][AGEN_mainform.Col_y] = AGEN_mainform.NA_y;
                                                    AGEN_mainform.Data_table_blocks.Rows[i]["SCALE"] = 1;
                                                    AGEN_mainform.Data_table_blocks.Rows[i]["BLOCK_NAME"] = AGEN_mainform.NA_name;
                                                    dataGridView_north_arrow_blocks.DataSource = AGEN_mainform.Data_table_blocks;
                                                    dataGridView_north_arrow_blocks.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);


                                                    Exista = true;

                                                }
                                            }

                                        }

                                        if (Exista == false)
                                        {

                                            AGEN_mainform.Data_table_blocks.Rows.Add();
                                            AGEN_mainform.Data_table_blocks.Rows[AGEN_mainform.Data_table_blocks.Rows.Count - 1]["TYPE"] = comboBox_type_of_block.Text;
                                            AGEN_mainform.Data_table_blocks.Rows[AGEN_mainform.Data_table_blocks.Rows.Count - 1][AGEN_mainform.Col_x] = AGEN_mainform.NA_x;
                                            AGEN_mainform.Data_table_blocks.Rows[AGEN_mainform.Data_table_blocks.Rows.Count - 1][AGEN_mainform.Col_y] = AGEN_mainform.NA_y;
                                            AGEN_mainform.Data_table_blocks.Rows[AGEN_mainform.Data_table_blocks.Rows.Count - 1]["SCALE"] = 1;
                                            AGEN_mainform.Data_table_blocks.Rows[AGEN_mainform.Data_table_blocks.Rows.Count - 1]["BLOCK_NAME"] = AGEN_mainform.NA_name;
                                            dataGridView_north_arrow_blocks.DataSource = AGEN_mainform.Data_table_blocks;
                                            dataGridView_north_arrow_blocks.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

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
                    Freeze_operations = false;

                    AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(false);

                }
                Ag.WindowState = FormWindowState.Normal;
            }
        }

        private void comboBox_type_of_block_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    string strTemplatePath = get_template_name_from_text_box();

                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                    if (System.IO.File.Exists(strTemplatePath) == false)
                    {
                        MessageBox.Show("template file not found");
                        Freeze_operations = false;
                        return;
                    }

                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {
                            AGEN_mainform.Template_is_open = true;
                            ThisDrawing = Doc;
                            DocumentManager1.CurrentDocument = ThisDrawing;
                            Functions.Incarca_existing_Blocks_to_combobox(comboBox_blocks_NA);
                        }
                    }

                    if (AGEN_mainform.Template_is_open == false)
                    {
                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                        Functions.Incarca_existing_Blocks_to_combobox(comboBox_blocks_NA);
                        AGEN_mainform.Template_is_open = true;
                    }

                    if (comboBox_blocks_NA.Items.Count > 1) comboBox_blocks_NA.Items.Insert(1, AGEN_mainform.insertNAtoMS);
                    if (comboBox_blocks_NA.Items.Count == 0)
                    {
                        comboBox_blocks_NA.Items.Add("");
                        comboBox_blocks_NA.Items.Add(AGEN_mainform.insertNAtoMS);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }


        }

        private void button_browser_dwt_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "autocad template files (*.dwt)|*.dwt";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_template_name.Text = fbd.FileName;
                }
            }
        }

        private void comboBox_dwgunits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[0].ToString())
            {
                AGEN_mainform.units_of_measurement = "f";
            }
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[1].ToString())
            {
                AGEN_mainform.units_of_measurement = "m";
            }
        }

        private void comboBox_units_precision_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_units_precision.Text == "0")
            {
                AGEN_mainform.round1 = 0;
            }
            else if (comboBox_units_precision.Text == "0.0")
            {
                AGEN_mainform.round1 = 1;
            }
            else if (comboBox_units_precision.Text == "0.00")
            {
                AGEN_mainform.round1 = 2;
            }
            else if (comboBox_units_precision.Text == "0.000")
            {
                AGEN_mainform.round1 = 3;
            }
            else
            {
                AGEN_mainform.round1 = 0;
            }
        }

        private void see_what_bands_are_picked()
        {
            if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
            {
                for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                {
                    if (AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                    {
                        if (Convert.ToString(AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]) == comboBox_bands_target_areas.Items[1].ToString())
                        {
                            is_main_view_picked = true;
                            i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                        }
                    }
                }

                if (is_main_view_picked == false)

                {
                    AGEN_mainform.Data_Table_regular_bands = Functions.creeaza_regular_band_data_table_structure();
                    AGEN_mainform.Data_Table_custom_bands = Functions.creeaza_custom_band_data_table_structure();
                }
            }
        }

        private void button_define_bands_Click(object sender, EventArgs e)
        {

            AGEN_mainform Ag = this.MdiParent as AGEN_mainform;
            if (Ag != null)
            {
                AGEN_mainform.locatie_config_file = Ag.get_textBox_config_file_location();

                if (System.IO.File.Exists(AGEN_mainform.locatie_config_file) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {



                    double x1 = 0;
                    double y1 = 0;
                    double x2 = 0;
                    double y2 = 0;

                    string strTemplatePath = AGEN_mainform.tpage_viewport_settings.get_template_name_from_text_box();

                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                    if (comboBox_bands_target_areas.Text == "")
                    {
                        MessageBox.Show("no type of viewport specified");
                        Freeze_operations = false;
                        return;
                    }
                    if (System.IO.File.Exists(strTemplatePath) == false)
                    {
                        MessageBox.Show("template file not found");
                        Freeze_operations = false;
                        return;
                    }

                    AGEN_mainform.Template_is_open = false;
                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {
                            AGEN_mainform.Template_is_open = true;
                            ThisDrawing = Doc;
                            DocumentManager1.CurrentDocument = ThisDrawing;
                            Functions.Incarca_existing_Blocks_to_combobox(comboBox_blocks_NA);

                            if (comboBox_blocks_NA.Items.Count > 1) comboBox_blocks_NA.Items.Insert(1, AGEN_mainform.insertNAtoMS);
                            if (comboBox_blocks_NA.Items.Count == 0)
                            {
                                comboBox_blocks_NA.Items.Add("");
                                comboBox_blocks_NA.Items.Add(AGEN_mainform.insertNAtoMS);
                            }

                        }

                    }

                    if (AGEN_mainform.Template_is_open == false)
                    {
                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);
                        Functions.Incarca_existing_Blocks_to_combobox(comboBox_blocks_NA);
                        if (comboBox_blocks_NA.Items.Count > 1) comboBox_blocks_NA.Items.Insert(1, AGEN_mainform.insertNAtoMS);
                        if (comboBox_blocks_NA.Items.Count == 0)
                        {
                            comboBox_blocks_NA.Items.Add("");
                            comboBox_blocks_NA.Items.Add(AGEN_mainform.insertNAtoMS);
                        }
                        AGEN_mainform.Template_is_open = true;
                    }


                    string Scale1 = AGEN_mainform.tpage_viewport_settings.Get_combobox_viewport_scale_text();
                    if (Scale1.Contains(":") == true)
                    {
                        Scale1 = Scale1.Substring(2, Scale1.Length - 2);
                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            AGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                        }
                    }


                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    Ag.WindowState = FormWindowState.Minimized;
                    see_what_bands_are_picked();


                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            #region main viewport
                            if (comboBox_bands_target_areas.Text == comboBox_bands_target_areas.Items[1].ToString())
                            {

                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nspecify the lower left point of the plan view");
                                PP1.AllowNone = false;
                                Point_res1 = Editor1.GetPoint(PP1);


                                if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Freeze_operations = false;
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;

                                Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                                Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\npick top right corner of the plan view");

                                if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                {
                                    Freeze_operations = false;
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

                                AGEN_mainform.Band_Separation = Math.Ceiling(3 * (Math.Abs(y2 - y1)) / 10) * 10;


                                string main_viewport_string = comboBox_bands_target_areas.Text;


                                if (is_main_view_picked == true)
                                {
                                    if (main_viewport_string != "")
                                    {

                                        if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (main_viewport_string == CT)
                                                {
                                                    AGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                                    AGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                                    AGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                                    AGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = AGEN_mainform.Vw_scale;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = main_viewport_string;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_h"] = Math.Abs(y2 - y1);
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["msx"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Top_y"] = y2;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = AGEN_mainform.Band_Separation;
                                                   

                                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;


                                                }
                                            }

                                        }

                                    }

                                }

                                else
                                {


                                    AGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                    AGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                    AGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                    AGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                    AGEN_mainform.Data_Table_regular_bands.Rows.Add();

                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = AGEN_mainform.Vw_scale;
                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = main_viewport_string;
                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_h"] = Math.Abs(y2 - y1);
                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["msx"] = AGEN_mainform.x_bands;
                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Top_y"] = y2;
                                    AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_separation"] = AGEN_mainform.Band_Separation;

                                    is_main_view_picked = true;
                                    AGEN_mainform.Exista_viewport_main = true;
                                }
                            }
                            #endregion

                            #region profile viewport

                            if (is_main_view_picked == true)
                            {

                                if (comboBox_bands_target_areas.Text == comboBox_bands_target_areas.Items[2].ToString())
                                {



                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                    Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of profile:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                    Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of profile:");

                                    if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
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





                                    string profile_band_string = comboBox_bands_target_areas.Text;
                                    bool Exista = false;

                                    if (profile_band_string != "")
                                    {

                                        if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (profile_band_string == CT)
                                                {
                                                    AGEN_mainform.Vw_prof_height = Math.Abs(y1 - y2);
                                                    AGEN_mainform.Vw_prof_y = y2;

                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = AGEN_mainform.Vw_scale;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = profile_band_string;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_h"] = Math.Abs(y2 - y1);
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["msx"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Top_y"] = y2;


                                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                                    Exista = true;

                                                }
                                            }

                                        }



                                        if (Exista == false)
                                        {

                                            AGEN_mainform.Vw_prof_height = Math.Abs(y1 - y2);

                                            AGEN_mainform.Vw_prof_y = y2;


                                            AGEN_mainform.Data_Table_regular_bands.Rows.Add();

                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = AGEN_mainform.Vw_scale;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = profile_band_string;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_h"] = Math.Abs(y2 - y1);
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["msx"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Top_y"] = y2;



                                            AGEN_mainform.Exista_viewport_prof = true;

                                        }



                                    }
                                }
                                #endregion

                                #region property viewport

                                if (comboBox_bands_target_areas.Text == comboBox_bands_target_areas.Items[3].ToString())
                                {

                                    if (AGEN_mainform.Vw_width == 0)
                                    {
                                        MessageBox.Show("first pick the plan view \r\naborted");
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                    Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of ownership band:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                    Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of ownership band:");

                                    if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
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

                                    AGEN_mainform.Point0_prop = new Point3d(AGEN_mainform.x_bands, Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2), 0);

                                    string property_band_string = comboBox_bands_target_areas.Text;
                                    bool Exista = false;

                                    if (property_band_string != "")
                                    {

                                        if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (property_band_string == CT)
                                                {

                                                    AGEN_mainform.Vw_prop_height = Math.Abs(y1 - y2);
                                                    AGEN_mainform.Vw_prop_y = y2;

                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = property_band_string;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_h"] = Math.Abs(y2 - y1);
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["msx"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Top_y"] = y2;
                                                   
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["x0"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);
                                                   
                                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                                    Exista = true;

                                                }
                                            }

                                        }

                                        if (Exista == false)
                                        {
                                            AGEN_mainform.Vw_prop_height = Math.Abs(y1 - y2);
                                            AGEN_mainform.Vw_prop_y = y2;

                                            AGEN_mainform.Data_Table_regular_bands.Rows.Add();

                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = 1;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = property_band_string;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_h"] = Math.Abs(y2 - y1);
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["msx"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Top_y"] = y2;
                                           
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["x0"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);

                                            AGEN_mainform.Exista_viewport_owner = true;

                                        }

                                    }
                                }
                                #endregion


                                #region crossing viewport

                                if (comboBox_bands_target_areas.Text == comboBox_bands_target_areas.Items[4].ToString())
                                {

                                    if (AGEN_mainform.Vw_width == 0)
                                    {
                                        MessageBox.Show("first pick the plan view \r\naborted");
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                    Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of crossing band:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                    Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of crossing band:");

                                    if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }


                                    y1 = Point_res1.Value.Y;
                                    y2 = Point_res2.Value.Y;

                                    AGEN_mainform.Point0_cross = new Point3d(AGEN_mainform.x_bands, Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2), 0);

                                    string crossing_band_string = comboBox_bands_target_areas.Text;
                                    bool Exista = false;

                                    if (crossing_band_string != "")
                                    {

                                        if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (crossing_band_string == CT)
                                                {
                                                    AGEN_mainform.Vw_cross_height = Math.Abs(y1 - y2);

                                                    AGEN_mainform.Vw_cross_y = y2;

                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = 1;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = crossing_band_string;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_h"] = Math.Abs(y2 - y1);
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["msx"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Top_y"] = y2;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["x0"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);

                                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                                    Exista = true;

                                                }
                                            }

                                        }

                                        if (Exista == false)
                                        {
                                            AGEN_mainform.Vw_cross_height = Math.Abs(y1 - y2);

                                            AGEN_mainform.Vw_cross_y = y2;

                                            AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = 1;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = crossing_band_string;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_h"] = Math.Abs(y2 - y1);
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["msx"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Top_y"] = y2;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["x0"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);

                                            AGEN_mainform.Exista_viewport_cross = true;
                                        }
                                    }
                                }
                                #endregion

                                #region material viewport

                                if (comboBox_bands_target_areas.Text == comboBox_bands_target_areas.Items[5].ToString())
                                {

                                    if (AGEN_mainform.Vw_width == 0)
                                    {
                                        MessageBox.Show("first pick the plan view \r\naborted");
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                    Functions.make_first_layout_active(Trans1, ThisDrawing.Database);

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify bottom of material band:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);


                                    if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }

                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                                    Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points_on_line();
                                    Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\nSpecify top of material band:");

                                    if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                                    {
                                        Freeze_operations = false;
                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        return;
                                    }


                                    y1 = Point_res1.Value.Y;
                                    y2 = Point_res2.Value.Y;

                                    AGEN_mainform.Point0_mat = new Point3d(AGEN_mainform.x_bands, Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2), 0);

                                    string material_band_string = comboBox_bands_target_areas.Text;
                                    bool Exista = false;

                                    if (material_band_string != "")
                                    {

                                        if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                        {

                                            for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                            {
                                                string CT = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                                if (material_band_string == CT)
                                                {
                                                    AGEN_mainform.Vw_mat_height = Math.Abs(y1 - y2);
                                                    AGEN_mainform.Vw_mat_y = y2;

                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = AGEN_mainform.Vw_scale;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = material_band_string;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_h"] = Math.Abs(y2 - y1);
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["msx"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["Top_y"] = y2;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["x0"] = AGEN_mainform.x_bands;
                                                    AGEN_mainform.Data_Table_regular_bands.Rows[i]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);

                                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                                    Exista = true;

                                                }
                                            }
                                        }

                                        if (Exista == false)
                                        {
                                            AGEN_mainform.Vw_mat_height = Math.Abs(y1 - y2);

                                            AGEN_mainform.Vw_mat_y = y2;


                                            AGEN_mainform.Data_Table_regular_bands.Rows.Add();
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = 1;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = material_band_string;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_h"] = Math.Abs(y2 - y1);
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["msx"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Top_y"] = y2;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["x0"] = AGEN_mainform.x_bands;
                                            AGEN_mainform.Data_Table_regular_bands.Rows[AGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["y0"] = Functions.calculate_vp_ms_y(AGEN_mainform.Band_Separation, y2);

                                            AGEN_mainform.Exista_viewport_mat = true;

                                        }

                                    }
                                }
                                #endregion

                            }
                            else
                            {
                                MessageBox.Show("pick first the main viewport");
                                Freeze_operations = false;
                                Ag.WindowState = FormWindowState.Normal;
                                return;
                            }

                            creeaza_display_data_table();
                            Trans1.Commit();
                        }
                    }

                    //ThisDrawing.CloseAndDiscard();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;



                AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(false);

                Ag.WindowState = FormWindowState.Normal;
            }
        }

        public void creeaza_display_data_table()
        {
            AGEN_mainform.Data_Table_display_bands = new System.Data.DataTable();

            AGEN_mainform.Data_Table_display_bands.Columns.Add("Band Name", typeof(string));
            AGEN_mainform.Data_Table_display_bands.Columns.Add("Custom Scale", typeof(double));

            if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
            {
                for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                {
                    if (AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value && AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] != DBNull.Value)
                    {

                        string bn = Convert.ToString(AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                        double cs = Convert.ToDouble(AGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"]);


                        AGEN_mainform.Data_Table_display_bands.Rows.Add();
                        AGEN_mainform.Data_Table_display_bands.Rows[AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                        AGEN_mainform.Data_Table_display_bands.Rows[AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Custom Scale"] = cs;

                    }
                }

                if (AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        if (AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                        {

                            string bn = Convert.ToString(AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);

                            AGEN_mainform.Data_Table_display_bands.Rows.Add();
                            AGEN_mainform.Data_Table_display_bands.Rows[AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                            AGEN_mainform.Data_Table_display_bands.Rows[AGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Custom Scale"] = 1;

                        }
                    }
                }

            }

            dataGridView_bands.DataSource = AGEN_mainform.Data_Table_display_bands;
            dataGridView_bands.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);



        }

        private void button_align_config_saveall_Click(object sender, EventArgs e)
        {


            AGEN_mainform Ag = this.MdiParent as AGEN_mainform;
            if (Ag != null)
            {
                AGEN_mainform.locatie_config_file = Ag.get_textBox_config_file_location();

                if (System.IO.File.Exists(AGEN_mainform.locatie_config_file) == true)
                {
                    AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(true);

                }

            }



        }

        private void TextBox_keypress_only_pozitive_integers(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_integer_pozitive_at_keypress(sender, e);
        }





        private void button_add_custom_Click(object sender, EventArgs e)
        {
            if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
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


        private void button_remove_regular_Click(object sender, EventArgs e)
        {
            if (comboBox_bands_target_areas.SelectedIndex < 1)
            {
                return;
            }

            string combo_val = "notselected";

            if (comboBox_bands_target_areas.SelectedIndex > 0)
            {
                combo_val = comboBox_bands_target_areas.Items[comboBox_bands_target_areas.SelectedIndex].ToString();
            }


            if (comboBox_bands_target_areas.SelectedIndex == 1) // main viewport - delete all
            {



                if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {


                    AGEN_mainform Ag = this.MdiParent as AGEN_mainform;

                    AGEN_mainform.locatie_config_file = Ag.get_textBox_config_file_location();

                    if (System.IO.File.Exists(AGEN_mainform.locatie_config_file) == true)
                    {
                        Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                        if (Excel1 == null)
                        {
                            MessageBox.Show("PROBLEM WITH EXCEL!");
                            return;
                        }

                        Excel1.Visible = AGEN_mainform.ExcelVisible;

                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(AGEN_mainform.locatie_config_file);

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                        try
                        {



                            //mainVP
                            W1.Range["B10"].Value = 0;
                            W1.Range["B11"].Value = 0;
                            W1.Range["B12"].Value = 0;
                            W1.Range["B13"].Value = 0;
                            W1.Range["B36"].Value = "False";
                            AGEN_mainform.Exista_viewport_main = false;
                            is_main_view_picked = false;


                            //crossing
                            W1.Range["B32"].Value = 0;
                            W1.Range["B33"].Value = 0;
                            W1.Range["B34"].Value = 0;
                            W1.Range["B35"].Value = 0;
                            W1.Range["B37"].Value = "False";
                            AGEN_mainform.Exista_viewport_cross = false;

                            //ownership
                            W1.Range["B28"].Value = 0;
                            W1.Range["B29"].Value = 0;
                            W1.Range["B30"].Value = 0;
                            W1.Range["B31"].Value = 0;
                            W1.Range["B38"].Value = "False";
                            AGEN_mainform.Exista_viewport_owner = false;

                            //profile
                            W1.Range["B25"].Value = 0;
                            W1.Range["B26"].Value = 0;
                            W1.Range["B27"].Value = 0;
                            W1.Range["B39"].Value = "False";
                            AGEN_mainform.Exista_viewport_prof = false;

                            //material
                            W1.Range["B41"].Value = "False";
                            W1.Range["B42"].Value = 0;
                            W1.Range["B43"].Value = 0;
                            W1.Range["B44"].Value = 0;
                            W1.Range["B45"].Value = 0;
                            AGEN_mainform.Exista_viewport_mat = false;

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

                            if (W_reg != null) W_reg.Delete();
                            if (W_reg != null) W_cust.Delete();

                            Workbook1.Save();
                            Workbook1.Close();
                            Excel1.Quit();
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);

                        }
                        finally
                        {
                            if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

                        }





                    }




                    AGEN_mainform.Data_Table_regular_bands = Functions.creeaza_regular_band_data_table_structure();
                    AGEN_mainform.Data_Table_custom_bands = Functions.creeaza_custom_band_data_table_structure();



                }
            }

            if (comboBox_bands_target_areas.SelectedIndex > 1) // viewport - delete one
            {
                if (AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {




                    AGEN_mainform Ag = this.MdiParent as AGEN_mainform;
                    AGEN_mainform.locatie_config_file = Ag.get_textBox_config_file_location();

                    if (System.IO.File.Exists(AGEN_mainform.locatie_config_file) == true)
                    {

                        for (int i = 0; i < AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                        {
                            if (AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                            {
                                string bn = AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                if (bn == combo_val)
                                {
                                    AGEN_mainform.Data_Table_regular_bands.Rows[i].Delete();
                                    i = AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                }
                            }
                        }


                        Microsoft.Office.Interop.Excel.Application Excel1 = new Microsoft.Office.Interop.Excel.Application();
                        if (Excel1 == null)
                        {
                            MessageBox.Show("PROBLEM WITH EXCEL!");
                            return;
                        }

                        Excel1.Visible = AGEN_mainform.ExcelVisible;


                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(AGEN_mainform.locatie_config_file);

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                        try
                        {
                            if (comboBox_bands_target_areas.SelectedIndex == 2)
                            {
                                //profile
                                W1.Range["B25"].Value = 0;
                                W1.Range["B26"].Value = 0;
                                W1.Range["B27"].Value = 0;
                                W1.Range["B39"].Value = "False";
                                AGEN_mainform.Exista_viewport_prof = false;
                            }

                            else if (comboBox_bands_target_areas.SelectedIndex == 3)
                            {
                                //owner
                                W1.Range["B28"].Value = 0;
                                W1.Range["B29"].Value = 0;
                                W1.Range["B30"].Value = 0;
                                W1.Range["B31"].Value = 0;
                                W1.Range["B38"].Value = "False";
                                AGEN_mainform.Exista_viewport_owner = false;
                            }
                            else if (comboBox_bands_target_areas.SelectedIndex == 4)
                            {
                                //crossing
                                W1.Range["B32"].Value = 0;
                                W1.Range["B33"].Value = 0;
                                W1.Range["B34"].Value = 0;
                                W1.Range["B35"].Value = 0;
                                W1.Range["B37"].Value = "False";
                                AGEN_mainform.Exista_viewport_cross = false;
                            }
                            else if (comboBox_bands_target_areas.SelectedIndex == 5)
                            {
                                //material
                                W1.Range["B41"].Value = "False";
                                W1.Range["B42"].Value = 0;
                                W1.Range["B43"].Value = 0;
                                W1.Range["B44"].Value = 0;
                                W1.Range["B45"].Value = 0;
                                AGEN_mainform.Exista_viewport_mat = false;
                            }


                            AGEN_mainform.tpage_setup.transfera_regular_band_to_excel(Workbook1);
                            AGEN_mainform.tpage_setup.transfera_custom_band_to_excel(Workbook1);

                            Workbook1.Save();
                            Workbook1.Close();
                            Excel1.Quit();
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);

                        }
                        finally
                        {
                            if (W1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W1);
                            if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                            if (Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

                        }

                    }







                    creeaza_display_data_table();



                }
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

    }
}
