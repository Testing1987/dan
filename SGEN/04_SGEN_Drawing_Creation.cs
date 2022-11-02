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
    public partial class SGEN_Drawing_Creation : Form
    {
        string col_scale = "Scale";
        string col_scaleName = "ScaleName";

        System.Data.DataTable Display_dt;
        private ContextMenuStrip ContextMenuStrip_open_alignment;

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(button_cut_sheets);
            lista_butoane.Add(button_output_location);
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(comboBox_end);

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
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(comboBox_end);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public SGEN_Drawing_Creation()
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
                Display_dt = Creaza_display_datatable_structure();
            }
            else if (Display_dt.Rows.Count == 0)
            {
                Display_dt = Creaza_display_datatable_structure();
            }
            string Col_dwg_name = _SGEN_mainform.Col_dwg_name;



            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = true;
                fbd.Filter = "alignment sheet (*.dwg)|*.dwg";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    for (int i = 0; i < fbd.FileNames.Count(); ++i)
                    {
                        string File1 = fbd.FileNames[i];
                        bool add_dwg_to_display = true;

                        if (Display_dt.Rows.Count > 0)
                        {
                            for (int k = 0; k < Display_dt.Rows.Count; ++k)
                            {
                                if (Display_dt.Rows[k][Col_dwg_name].ToString() == File1)
                                {
                                    add_dwg_to_display = false;
                                    k = Display_dt.Rows.Count;
                                }
                            }
                        }

                        if (_SGEN_mainform.dt_sheet_index != null && _SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                        {

                            bool is_found = false;
                            for (int k = 0; k < _SGEN_mainform.dt_sheet_index.Rows.Count; ++k)
                            {
                                if (_SGEN_mainform.dt_sheet_index.Rows[k][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                                {
                                    string nume_sheet = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[k][_SGEN_mainform.Col_dwg_name]);
                                    string nume_selected = System.IO.Path.GetFileNameWithoutExtension(File1);
                                    if (nume_selected.ToLower() == nume_sheet.ToLower())
                                    {
                                        is_found = true;
                                        k = _SGEN_mainform.dt_sheet_index.Rows.Count;
                                    }
                                }
                            }

                            if (is_found == false) add_dwg_to_display = false;

                            if (add_dwg_to_display == true)
                            {
                                Display_dt.Rows.Add();
                                Display_dt.Rows[Display_dt.Rows.Count - 1][Col_dwg_name] = File1;
                            }
                        }
                    }
                }
            }


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
                Display_dt = Creaza_display_datatable_structure();
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
                        string fisier_generat = Display_dt.Rows[Index1][_SGEN_mainform.Col_dwg_name].ToString();
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

        private void button_output_location_Click(object sender, System.EventArgs e)
        {
            try
            {
                _SGEN_mainform.output_folder = _SGEN_mainform.tpage_settings.get_textbox_autpot_content();
                Process.Start(@"" + _SGEN_mainform.output_folder);
            }
            catch (System.Exception)
            {

                MessageBox.Show("Please specify the output folder in the project settings page!");
            }
        }
        private void button_load_dwgs_in_comboboxs1_Click(object sender, EventArgs e)
        {
            comboBox_start.Items.Clear();
            comboBox_start.Items.Add("");



            _SGEN_mainform.tpage_sheetindex.Build_sheet_index_dt_from_excel();


            if (_SGEN_mainform.dt_sheet_index != null && _SGEN_mainform.dt_sheet_index.Rows.Count > 0)
            {
                for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                {
                    if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                    {
                        comboBox_start.Items.Add(Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name]));

                    }
                }
            }
        }

        private void button_load_dwgs_in_comboboxs2_Click(object sender, EventArgs e)
        {
            comboBox_end.Items.Clear();
            comboBox_end.Items.Add("");

            _SGEN_mainform.tpage_sheetindex.Build_sheet_index_dt_from_excel();

            if (_SGEN_mainform.dt_sheet_index != null && _SGEN_mainform.dt_sheet_index.Rows.Count > 0)
            {
                for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                {
                    if (_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                    {
                        comboBox_end.Items.Add(Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name]));
                    }
                }
            }
        }

        private void button_cut_sheets_Click(object sender, EventArgs e)
        {


            DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            if (DocumentManager1.Count == 0)
            {
                string strTemplatePath = "acad.dwt";
                Document acDoc = DocumentManager1.Add(strTemplatePath);
                DocumentManager1.MdiActiveDocument = acDoc;
            }




            if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }


            if (_SGEN_mainform.Vw_height <= 0 || _SGEN_mainform.Vw_height <= 0)
            {
                MessageBox.Show("no main viewport dimension specified\r\nOperation aborted");

                return;
            }





            if (_SGEN_mainform.dt_sheet_index == null || _SGEN_mainform.dt_sheet_index.Rows.Count == 0)
            {
                MessageBox.Show("no sheet index data found\r\nOperation aborted");
                return;
            }


            try
            {
                try
                {

                    set_enable_false();
                    System.Data.DataTable Dtp = null;


                    string ProjF = _SGEN_mainform.project_main_folder;
                    if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                    {
                        ProjF = ProjF + "\\";
                    }

                    if (System.IO.Directory.Exists(ProjF) == true)
                    {

                    }
                    else
                    {
                        MessageBox.Show("there is no such a project folder");
                        set_enable_true();
                        return;
                    }

                    if (_SGEN_mainform.dt_sheet_index.Rows.Count == 0)
                    {
                        MessageBox.Show("sheet index table is empty\r\nOperation aborted");
                        set_enable_true();
                        return;
                    }

                    string Template_file_name = _SGEN_mainform.tpage_settings.get_template_name_from_text_box();
                    string Output_folder = _SGEN_mainform.tpage_settings.get_output_folder_from_text_box();

                    if (Output_folder.Substring(Output_folder.Length - 1, 1) != "\\")
                    {
                        Output_folder = Output_folder + "\\";
                    }

                    Point3d ms_point = new Point3d();
                    Point3d ps_point_plan_view = new Point3d(_SGEN_mainform.Vw_ps_x, _SGEN_mainform.Vw_ps_y, 0);

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

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord Btrecord = Functions.get_modelspace(Trans1, ThisDrawing.Database);

                            //Btrecord.UpgradeOpen();






                            #region creaza new file

                            if (Creaza_new_file == true)
                            {

                                List<int> lista_generation = new List<int>();

                                if (comboBox_start.Text != "" & comboBox_end.Text != "")
                                {
                                    lista_generation = create_band_list_of_dwg(comboBox_start.Text, comboBox_end.Text);
                                }



                                if (lista_generation.Count == 0)
                                {
                                    for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                    {
                                        lista_generation.Add(i);
                                    }
                                }

                                Document New_doc = DocumentCollectionExtension.Add(DocumentManager1, Template_file_name);
                                DocumentManager1.MdiActiveDocument = New_doc;

                                string seg_name = _SGEN_mainform.tpage_settings.get_combobox_segment_name_value();
                                if (seg_name == "")
                                {
                                    seg_name = _SGEN_mainform.tpage_settings.get_textBox_client_name();
                                }

                                if (seg_name == "")
                                {
                                    seg_name = _SGEN_mainform.tpage_settings.get_textBox_project_name();
                                }
                                if (seg_name == "")
                                {
                                    seg_name = _SGEN_mainform.tpage_settings.get_textBox_prefix_name();
                                }
                                if (seg_name == "")
                                {
                                    seg_name = "SGEN_ALL_sheets";
                                }
                                string fname0 = Output_folder + _SGEN_mainform.dt_sheet_index.Rows[lista_generation[0]][_SGEN_mainform.Col_dwg_name].ToString() + ".dwg";
                                if (checkBox_multiple_layouts.Checked == true)
                                {
                                    fname0 = Output_folder + seg_name + ".dwg";
                                }


                                if (System.IO.File.Exists(fname0) == true)
                                {
                                    MessageBox.Show(fname0 + " already exists.\r\noperation aborted");
                                    Display_dt = Creaza_display_datatable_structure();
                                    set_enable_true();
                                    return;
                                }

                                Display_dt = Creaza_display_datatable_structure();

                                using (DocumentLock lock2 = New_doc.LockDocument())
                                {

                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = New_doc.Database.TransactionManager.StartTransaction())
                                    {
                                        if (_SGEN_mainform.dt_sheet_index.Rows[0][_SGEN_mainform.Col_dwg_name] == DBNull.Value)
                                        {
                                            MessageBox.Show("no sheet index dwg name specified");
                                            set_enable_true();
                                            return;
                                        }
                                        string nume1 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[0]][_SGEN_mainform.Col_dwg_name]);


                                        BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, New_doc.Database);
                                        BtrecordPS.UpgradeOpen();
                                        Layout layout1 = Functions.get_first_layout(Trans2, New_doc.Database);
                                        layout1.UpgradeOpen();

                                        if (checkBox_multiple_layouts.Checked == true) nume1 = "Sgen1234";

                                        layout1.LayoutName = nume1;

                                        if (checkBox_multiple_layouts.Checked == true)
                                        {
                                            LayoutManager LayoutManager1 = (LayoutManager)Autodesk.AutoCAD.DatabaseServices.LayoutManager.Current;
                                            Functions.Creaza_layer(_SGEN_mainform.Layer_name_Main_Viewport, 4, false);
                                            for (int i = lista_generation.Count - 1; i >= 0; --i)
                                            {
                                                string nume2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_dwg_name]);
                                                double scale2 = _SGEN_mainform.Vw_scale;
                                                string scalename2 = "";
                                                if (_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scale] != DBNull.Value)
                                                {
                                                    scale2 = Convert.ToDouble(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scale]);
                                                }

                                                if (_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scaleName] != DBNull.Value)
                                                {
                                                     scalename2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scaleName]).Replace("'", "");
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
                                                                    scale2 = nr2 / nrr2;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                                LayoutManager1.CopyLayout(nume1, nume2);
                                                LayoutManager1.CurrentLayout = nume2;
                                                Layout layout2 = Trans2.GetObject(LayoutManager1.GetLayoutId(nume2), OpenMode.ForRead) as Layout;
                                                if (layout2 != null)
                                                {

                                                    BlockTableRecord btr2 = Trans2.GetObject(layout2.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                                                    if (btr2 != null)
                                                    {
                                                        ms_point = new Point3d((double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_x], (double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_y], 0);
                                                        double Twist = 2 * Math.PI - (double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_rot] * Math.PI / 180;

                                                        #region plan view

                                                        ObjectContextManager ocm = New_doc.Database.ObjectContextManager;

                                                        ObjectContextCollection occ = null;
                                                        if (ocm != null)
                                                        {
                                                            occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                                                        }


                                                        Viewport Viewport_main = Functions.Create_viewport(ms_point, ps_point_plan_view, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, scale2, Twist);
                                                        Viewport_main.Layer = _SGEN_mainform.Layer_name_Main_Viewport;

                                                        btr2.AppendEntity(Viewport_main);
                                                        Trans2.AddNewlyCreatedDBObject(Viewport_main, true);
                                                        Viewport_main.On = true;


                                                        if (occ != null && scalename2 != "")
                                                        {
                                                            AnnotationScale Anno_scale = occ.GetContext(scalename2) as AnnotationScale;
                                                            if (Anno_scale != null)
                                                            {
                                                                Viewport_main.AnnotationScale = Anno_scale;
                                                            }

                                                        }



                                                        ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                                                        oBJiD_COL.Add(Viewport_main.ObjectId);
                                                        DrawOrderTable DrawOrderTable2 = Trans2.GetObject(btr2.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                                                        DrawOrderTable2.MoveToBottom(oBJiD_COL);

                                                        #endregion
                                                    }

                                                }

                                            }

                                            LayoutManager1.DeleteLayout(nume1);


                                        }

                                        Trans2.Commit();
                                        New_doc.Database.SaveAs(fname0, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);
                                    }
                                }



                                New_doc.CloseAndDiscard();

                                if (checkBox_multiple_layouts.Checked == false)
                                {
                                    if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_generation.Count; ++i)
                                        {
                                            string Fisier2 = Output_folder + _SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_dwg_name].ToString() + ".dwg";

                                            if (i > 0)
                                            {
                                                string Fisier1 = Output_folder + _SGEN_mainform.dt_sheet_index.Rows[lista_generation[i - 1]][_SGEN_mainform.Col_dwg_name].ToString() + ".dwg";
                                                System.IO.File.Copy(Fisier1, Fisier2, false);
                                            }

                                            Display_dt.Rows.Add();
                                            Display_dt.Rows[Display_dt.Rows.Count - 1][_SGEN_mainform.Col_dwg_name] = Fisier2;
                                        }
                                    }
                                    for (int i = 0; i < lista_generation.Count; ++i)
                                    {
                                        List<Polyline> lista_poly = new List<Polyline>();

                                        string dwg_name = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_dwg_name]);
                                        string Fisier = Output_folder + dwg_name + ".dwg";
                                        using (Database Database2 = new Database(false, true))
                                        {

                                            Database2.ReadDwgFile(Fisier, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                            //System.IO.FileShare.ReadWrite, false, null);
                                            Database2.CloseInput(true);

                                            HostApplicationServices.WorkingDatabase = Database2;
                                            Functions.Creaza_layer_on_database(Database2, _SGEN_mainform.Layer_name_Main_Viewport, 4, false);




                                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
                                            {

                                                Functions.make_first_layout_active(Trans2, Database2);

                                                BlockTableRecord BtrecordPS = Functions.get_first_layout_as_paperspace(Trans2, Database2);
                                                BtrecordPS.UpgradeOpen();
                                                Layout Layout1 = Functions.get_first_layout(Trans2, Database2);
                                                Layout1.UpgradeOpen();
                                                Layout1.LayoutName = _SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_dwg_name].ToString();







                                                ms_point = new Point3d((double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_x], (double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_y], 0);
                                                double Twist = 2 * Math.PI - (double)_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][_SGEN_mainform.Col_rot] * Math.PI / 180;

                                                double scale2 = _SGEN_mainform.Vw_scale;
                                                string scalename2 = "";



                                                if (_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scale] != DBNull.Value)
                                                {
                                                    scale2 = Convert.ToDouble(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scale]);
                                                }

                                                if (_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scaleName] != DBNull.Value)
                                                {
                                                    scalename2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[lista_generation[i]][col_scaleName]).Replace("'", "");
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
                                                                    scale2 = nr2 / nrr2;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }




                                                #region plan view

                                                ObjectContextManager ocm = Database2.ObjectContextManager;

                                                ObjectContextCollection occ = null;
                                                if (ocm != null)
                                                {
                                                    occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                                                }

                                                Viewport Viewport_main = Functions.Create_viewport(ms_point, ps_point_plan_view, _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, scale2, Twist);
                                                Viewport_main.Layer = _SGEN_mainform.Layer_name_Main_Viewport;

                                                BtrecordPS.AppendEntity(Viewport_main);
                                                Trans2.AddNewlyCreatedDBObject(Viewport_main, true);


                                                if (occ != null && scalename2 != "")
                                                {
                                                    AnnotationScale Anno_scale = occ.GetContext(scalename2) as AnnotationScale;
                                                    if (Anno_scale != null)
                                                    {
                                                        Viewport_main.AnnotationScale = Anno_scale;
                                                    }

                                                }


                                                ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                                                oBJiD_COL.Add(Viewport_main.ObjectId);
                                                DrawOrderTable DrawOrderTable2 = Trans2.GetObject(BtrecordPS.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                                                DrawOrderTable2.MoveToBottom(oBJiD_COL);

                                                #endregion





                                                Trans2.Commit();

                                                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                Database2.SaveAs(Fisier, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);







                                            }
                                        }
                                    }
                                }

                                if (checkBox_multiple_layouts.Checked == false)
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
                            #endregion


                            #region update file with new vp

                            if (Creaza_new_file == false)
                            {




                                for (int i = 0; i < Display_dt.Rows.Count; ++i)
                                {
                                    List<Polyline> lista_poly = new List<Polyline>();
                                    if (Display_dt.Rows[i][_SGEN_mainform.Col_dwg_name] != DBNull.Value)
                                    {
                                        string file1 = Display_dt.Rows[i][_SGEN_mainform.Col_dwg_name].ToString();
                                        string nume_fara_ext = System.IO.Path.GetFileNameWithoutExtension(file1);

                                        if (System.IO.File.Exists(file1) == true)
                                        {
                                            Point3d ms_pt = new Point3d();
                                            Point3d ps_pt = new Point3d();

                                            int si_index = -1;








                                            for (int j = 0; j < _SGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                            {
                                                string si_name = _SGEN_mainform.dt_sheet_index.Rows[j][_SGEN_mainform.Col_dwg_name].ToString();
                                                if (si_name.ToLower() == nume_fara_ext.ToLower())
                                                {
                                                    si_index = j;

                                                    j = _SGEN_mainform.dt_sheet_index.Rows.Count;
                                                }
                                            }



                                            if (si_index >= 0)
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
                                                        Layout1.LayoutName = _SGEN_mainform.dt_sheet_index.Rows[si_index][_SGEN_mainform.Col_dwg_name].ToString();


                                                        double scale2 = _SGEN_mainform.Vw_scale;

                                                        if (_SGEN_mainform.dt_sheet_index.Rows[si_index][col_scale] != DBNull.Value)
                                                        {
                                                            scale2 = Convert.ToDouble(_SGEN_mainform.dt_sheet_index.Rows[si_index][col_scale]);
                                                        }

                                                        if (_SGEN_mainform.dt_sheet_index.Rows[si_index][col_scaleName] != DBNull.Value)
                                                        {
                                                            string scalename2 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[si_index][col_scaleName]);
                                                            if (scalename2.Contains(":") == true)
                                                            {
                                                                string[] text_array1 = scalename2.Replace("'", "").Split(Convert.ToChar(":"));
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
                                                                            scale2 = nr2 / nrr2;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }



                                                        #region VP plan view


                                                        ms_pt = new Point3d((double)_SGEN_mainform.dt_sheet_index.Rows[si_index][_SGEN_mainform.Col_x], (double)_SGEN_mainform.dt_sheet_index.Rows[si_index][_SGEN_mainform.Col_y], 0);
                                                        double twist1 = 2 * Math.PI - (double)_SGEN_mainform.dt_sheet_index.Rows[si_index][_SGEN_mainform.Col_rot] * Math.PI / 180;

                                                        Creaza_viewport_on_alignment_on_existing(Database2, _SGEN_mainform.Layer_name_Main_Viewport, ms_pt,
                                                                                                                        new Point3d(_SGEN_mainform.Vw_ps_x, _SGEN_mainform.Vw_ps_y, 0),
                                                                                                                        _SGEN_mainform.Vw_width, _SGEN_mainform.Vw_height, scale2, twist1);

                                                        #endregion





                                                        Trans2.Commit();
                                                    }
                                                    HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                                                    Database2.SaveAs(file1, true, DwgVersion.Current, ThisDrawing.Database.SecurityParameters);







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

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
            catch (System.AccessViolationException ex1)
            {
                MessageBox.Show(ex1.Message);

            }


            set_enable_true();

        }



        private void Creaza_viewport_on_alignment_on_existing(Database Database2, string Layer_name_Viewport, Point3d MSpoint, Point3d PSpoint, double width1, double height1, double scale1, double twist1,string scalename2="")
        {
            HostApplicationServices.WorkingDatabase = Database2;
            Functions.Creaza_layer_on_database(Database2, Layer_name_Viewport, 4, false);

            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans2 = Database2.TransactionManager.StartTransaction())
            {
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

                ObjectContextManager ocm = Database2.ObjectContextManager;

                ObjectContextCollection occ = null;
                if (ocm != null)
                {
                    occ = ocm.GetContextCollection("ACDB_ANNOTATIONSCALES");
                }


                Viewport new_viewport = Functions.Create_viewport(MSpoint, PSpoint, width1, height1, scale1, twist1);
                new_viewport.Layer = Layer_name_Viewport;
                BtrecordPS.AppendEntity(new_viewport);
                Trans2.AddNewlyCreatedDBObject(new_viewport, true);

                if (occ != null && scalename2 != "")
                {
                    AnnotationScale Anno_scale = occ.GetContext(scalename2) as AnnotationScale;
                    if (Anno_scale != null)
                    {
                        new_viewport.AnnotationScale = Anno_scale;
                    }

                }

                ObjectIdCollection oBJiD_COL = new ObjectIdCollection();
                oBJiD_COL.Add(new_viewport.ObjectId);
                DrawOrderTable DrawOrderTable2 = Trans2.GetObject(BtrecordPS.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as DrawOrderTable;
                DrawOrderTable2.MoveToBottom(oBJiD_COL);

                Trans2.Commit();

            }
        }



        public List<int> create_band_list_of_dwg(string start1, string end1)
        {
            List<int> lista1 = new List<int>();
            if (_SGEN_mainform.dt_sheet_index != null)
            {
                if (_SGEN_mainform.dt_sheet_index.Rows.Count > 0)
                {
                    bool adauga = false;
                    for (int i = 0; i < _SGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                    {
                        string nume1 = Convert.ToString(_SGEN_mainform.dt_sheet_index.Rows[i][_SGEN_mainform.Col_dwg_name]);
                        if (nume1.ToUpper() == start1.ToUpper() || start1 == "" && end1 == "")
                        {
                            adauga = true;
                        }

                        if (adauga == true)
                        {
                            lista1.Add(i);
                        }
                        if (nume1.ToUpper() == end1.ToUpper())
                        {
                            adauga = false;
                        }

                    }
                }
            }
            return lista1;
        }



        public static System.Data.DataTable Creaza_display_datatable_structure()
        {

            string Col_dwg_name = _SGEN_mainform.Col_dwg_name;

            System.Type type_dwg_name = typeof(string);

            List<string> Lista1 = new List<string>();
            List<System.Type> Lista2 = new List<System.Type>();


            Lista1.Add(Col_dwg_name);

            Lista2.Add(type_dwg_name);



            System.Data.DataTable dt1 = new System.Data.DataTable();

            for (int i = 0; i < Lista1.Count; ++i)
            {
                dt1.Columns.Add(Lista1[i], Lista2[i]);
            }
            return dt1;
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


    }
}
