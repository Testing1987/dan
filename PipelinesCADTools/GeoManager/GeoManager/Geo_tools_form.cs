using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ImportExport;
using Autodesk.Gis.Map.ObjectData;
using Autodesk.Gis.Map.Project;

namespace Alignment_mdi
{
    public partial class Geo_tools_form : Form
    {


        string Tab = "\t";
        string vbcrlf = "\r\n";
        System.Data.DataTable Data_table_OD_attrib_existing;
        System.Data.DataTable Data_table_centerline;

        static public List<string> col_station_labels;

        string Col_x = "X";
        string Col_y = "Y";
        string Col_z = "Z";
        string cl_id_for_temp = null;
        string layer_no_plot = "NO PLOT";

        string project_type = "2d";

        List<string> col_labels_zoom;

        Point3d picked_pt = new Point3d(123.123, 123.123, 123.123);
        Point3d pt_on_poly = new Point3d(123.123, 123.123, 123.123);



        List<string> List_all_objId;
        List<ObjectId> List_update_objId;
        List<int> List_update_row_index;
        System.Data.DataTable Table_filter;


        string Object_id_grid_current;

        bool Freeze_operations = false;

        List<int> List_red;
        List<int> List_yellow;
        List<int> List_blue;
        List<string> List_red_objId;
        List<string> List_yellow_objId;
        List<string> List_blue_objId;
        bool Is_update_running = false;

        string Correct_table = "";
        string Correct_layer = "";
        MMGeoTools.Form_processing Processing_form;
        List<string> List_of_tables_on_layer;

        int Oldw1 = 148;
        int Oldl2 = 162;
        int Oldw2 = 148;
        int Oldll2 = 170;
        int Oldl3 = 316;

        //overlap form

        double Ultimul_top;
        double Spacing;
        ObjectId[] Empty_array;

        System.Data.DataTable dt_layer = null;

        public Geo_tools_form()
        {
            InitializeComponent();
            DataGridView_data.MultiSelect = true;
            col_labels_zoom = new List<string>();
            col_station_labels = new List<string>();

        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_export);
            lista_butoane.Add(button_load_layers);
            lista_butoane.Add(button_browse_select_output_folder);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_export);
            lista_butoane.Add(button_load_layers);
            lista_butoane.Add(button_browse_select_output_folder);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }



        private void OD_TABLE_form_Load(object sender, EventArgs e)
        {
            //Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("._blockicon" + "\r\n", true, false, false);
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();

            ToolTip1.SetToolTip(this.button_add_OD_table, "All Objects on Selected Layer Will Be Assigned the Selected Object Data Table." + "\r\nData from the Selected Table Will Be Maintained. All Other Tables and Data Will Be Purged.");
            ToolTip1.SetToolTip(this.button_Update_object_data, "Any Changes Made on Table Will Update Features in Drawing.");
            ToolTip1.SetToolTip(this.button_refresh_grid, "Refresh and Display All Features on the Selected Layer. Also Refreshes the Statistics Table.");
            ToolTip1.SetToolTip(this.button_zoom, "Zooms to Selected Row on Tables Corresponding Feature.");
            ToolTip1.SetToolTip(this.comboBox_od_existing_tables, "Specify Appropriate Data Table Based on Current Layer.");
            ToolTip1.SetToolTip(this.comboBox_layers_blocks_geomanager, "Specify Which Layer You Want to Load the Object Data Table.");
            ToolTip1.SetToolTip(this.button_refresh_layer_tables, "Load Layer Names & Object Data Tables from Current Drawing.");
            ToolTip1.SetToolTip(this.button_zoom_row_object_data, "Designate Geometry in Model Space You Want to Display on Table.");
            ToolTip1.SetToolTip(this.button_Filter, "Filter to Issues Identified in Statistics Table.");


            StopService("TabletInputService", 36000000);

            System.Diagnostics.Process[] proc = System.Diagnostics.Process.GetProcessesByName("tabtip");

            for (int i = 0; i < proc.Length; i = i + 1)
            {
                proc[i].Kill();
            }
            panel_excel.Visible = false;



            // overlap form
            Functions.Incarca_existing_layers_to_combobox(comboBox_GIC_Layer1);

            Ultimul_top = comboBox_GIC_Layer1.Top;
            Spacing = Convert.ToInt32(1.3 * comboBox_GIC_Layer1.Height);
            label_processing1.Visible = false;
            label_processing2.Visible = false;

            System.Windows.Forms.Control[] Del1;

            int Index1 = 0;
            Del1 = new System.Windows.Forms.Control[Index1 + 1];
            foreach (System.Windows.Forms.Control control1 in panel_layerpanel.Controls)
            {


                if (control1.Top > comboBox_GIC_Layer1.Top)
                {
                    Array.Resize(ref Del1, Index1 + 1);
                    Del1[Index1] = control1;
                    Index1 = Index1 + 1;
                }
            }

            if (Del1 == null)
            {
                for (int i = 0; i < Del1.Length; ++i)
                {
                    panel_layerpanel.Controls.Remove(Del1[i]);
                }
            }


        }

        public static void StopService(string serviceName, int timeoutMilliseconds)
        {
            System.ServiceProcess.ServiceController service = new System.ServiceProcess.ServiceController(serviceName);
            try
            {
                TimeSpan timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);

                service.Stop();
                service.WaitForStatus(System.ServiceProcess.ServiceControllerStatus.Stopped, timeout);
            }
            catch
            {
                // ...
            }
        }

        private void OD_TABLE_form_Resize(object sender, EventArgs e)
        {


            int FormH = 749;
            int FormW = 953;

            if (this.Height > FormH)
            {
                panel_grid.Height = this.Height - (FormH - 382);
                tabControl1.Height = this.Height - (FormH - 709);
                button_add_OD_table.Top = this.Height - (FormH - 638);
                panel_stats.Top = this.Height - (FormH - 498);
                button_Filter.Top = this.Height - (FormH - 638);
                panel_navigation.Top = this.Height - (FormH - 498);
            }
            if (this.Width > FormW)
            {
                panel_grid.Width = this.Width - (FormW - 909);
                tabControl1.Width = this.Width - (FormW - 933);
                panel_logo.Left = this.Width - (FormW - 788);
                panel_logo1.Left = this.Width - (FormW - 788);
                label_Apply_Changes.Left = this.Width - (FormW - 400);

            }


            label_ver.Location = new System.Drawing.Point(panel_navigation.Location.X + 2, button_add_OD_table.Location.Y);
            button_Update_object_data.Location = new System.Drawing.Point(panel_grid.Location.X + panel_grid.Size.Width - button_Update_object_data.Size.Width, button_add_OD_table.Location.Y);
            label_Apply_Changes.Location = new System.Drawing.Point(button_Update_object_data.Location.X - 144, panel_grid.Location.Y + panel_grid.Size.Height + 1);
        }

        private int GetMaxSize(List<string> List_of_items, int SW)
        {
            Graphics g = CreateGraphics();
            SizeF size;
            int oldSize = SW - 20;

            foreach (string item1 in List_of_items)
            {
                size = g.MeasureString(item1, comboBox_layers_blocks_geomanager.Font);

                if (size.Width > oldSize)
                {
                    oldSize = (int)size.Width;

                }
            }

            return oldSize + 20;
        }


        private void button_load_layers_and_data_tables_Click(object sender, EventArgs e)
        {

            {
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    Data_table_OD_attrib_existing = new System.Data.DataTable();
                    List_all_objId = new List<string>();

                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {

                            textBox_no_od_2.Text = "";
                            textBox_no_od_zero.Text = "";
                            textBox_no_rows.Text = "";
                            textBox_no_tables.Text = "";
                            textBox_no_wrong_od.Text = "";

                            DataGridView_data.DataSource = "";
                            DataGridView_data.Refresh();
                            comboBox_od_existing_tables.Items.Clear();



                            if (radioButton_OD.Checked == true)
                            {
                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                                Nume_tables = Tables1.GetTableNames();

                                for (int i = 0; i < Nume_tables.Count; i = i + 1)
                                {
                                    String Tabla1 = Nume_tables[i];
                                    if (comboBox_od_existing_tables.Items.Contains(Tabla1) == false)
                                    {
                                        comboBox_od_existing_tables.Items.Add(Tabla1);
                                    }
                                }

                                Functions.Incarca_existing_layers_to_combobox(comboBox_layers_blocks_geomanager);
                            }


                            if (radioButton_BLOCKS.Checked == true)
                            {
                                Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_layers_blocks_geomanager);
                            }


                            if (comboBox_layers_blocks_geomanager.Items.Count > 0)
                            {
                                comboBox_layers_blocks_geomanager.SelectedIndex = 0;
                            }

                            List<string> List1 = new List<string>();

                            for (int i = 0; i < comboBox_layers_blocks_geomanager.Items.Count; ++i)
                            {

                                string item1 = comboBox_layers_blocks_geomanager.Items[i].ToString();
                                List1.Add(item1);

                            }

                            List<string> List2 = new List<string>();

                            for (int i = 0; i < comboBox_od_existing_tables.Items.Count; ++i)
                            {

                                string item2 = comboBox_od_existing_tables.Items[i].ToString();
                                List2.Add(item2);

                            }

                            if (radioButton_OD.Checked == true)
                            {

                                int Neww1 = GetMaxSize(List1, Oldw1);
                                comboBox_layers_blocks_geomanager.Width = Neww1;

                                comboBox_od_existing_tables.Left = Oldl2 + (Neww1 - Oldw1);

                                label_correct_od_table.Left = Oldll2 + (Neww1 - Oldw1);

                                int Neww2 = GetMaxSize(List2, Oldw2);
                                comboBox_od_existing_tables.Width = Neww2;


                                button_refresh_grid.Left = Oldl3 + (Neww1 - Oldw1) + (Neww2 - Oldw2);
                            }

                            if (radioButton_BLOCKS.Checked == true)
                            {

                                int Neww1 = GetMaxSize(List1, Oldw1);
                                comboBox_layers_blocks_geomanager.Width = Neww1;

                                button_refresh_grid.Left = Oldl2 + (Neww1 - Oldw1);
                            }

                            label_drawing_name.Text = ThisDrawing.Database.OriginalFileName;
                            this.Refresh();
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }



        private void button_LOAD_DATA_Click(object sender, EventArgs e)
        {

            if (Freeze_operations == false)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
                {
                    MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                    label_processing1.Visible = false;
                    return;
                }

                bool Update_data = true;

                if (DataGridView_data.RowCount > 0)
                {
                    if (MessageBox.Show("You have not applied changes. Any changes made to the object data table will not be saved! Do you wish to continue?", "WARNING", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
                    {
                        Update_data = false;
                    }
                }


                if (radioButton_OD.Checked == true)
                {
                    if (this.comboBox_layers_blocks_geomanager.Text == "" | this.comboBox_od_existing_tables.Text == "")
                    {
                        MessageBox.Show("Please select a layer and an object data table!");
                        label_processing1.Visible = false;
                        return;
                    }

                    if (Update_data == true)
                    {
                        Freeze_operations = true;
                        try
                        {
                            DataGridView_data.DataSource = null;
                            DataGridView_data.Refresh();
                            textBox_no_rows.Text = "";
                            textBox_no_od_2.Text = "";
                            textBox_no_od_zero.Text = "";
                            textBox_no_tables.Text = "";
                            textBox_no_wrong_od.Text = "";

                            Data_table_OD_attrib_existing = new System.Data.DataTable();
                            Correct_table = comboBox_od_existing_tables.Text;
                            Correct_layer = comboBox_layers_blocks_geomanager.Text;

                            {
                                List_red = new List<int>();
                                List_yellow = new List<int>();
                                List_blue = new List<int>();
                                List_red_objId = new List<string>();
                                List_yellow_objId = new List<string>();
                                List_blue_objId = new List<string>();

                                List_all_objId = new List<string>();
                                List_update_objId = new List<ObjectId>();
                                List_update_row_index = new List<int>();


                                Data_table_OD_attrib_existing.Columns.Add("OBJECT_ID", typeof(ObjectId));
                                Data_table_OD_attrib_existing.Columns.Add("OD_TABLE_COUNT", typeof(int));
                                Data_table_OD_attrib_existing.Columns.Add("OBJECT_TYPE", typeof(String));
                                Data_table_OD_attrib_existing.Columns.Add("BLOCK_NAME", typeof(String));

                                List_of_tables_on_layer = new List<string>();

                                label_processing1.Visible = true;
                                this.Refresh();



                                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                                {
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;


                                        Autodesk.Gis.Map.ObjectData.Table Tabla1;

                                        if (Tables1.IsTableDefined(Correct_table) == true)
                                        {
                                            Tabla1 = Tables1[Correct_table];


                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            for (int i = 0; i < Field_defs1.Count; ++i)
                                            {
                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                string Nume_field = Field_def1.Name;


                                                if (Data_table_OD_attrib_existing.Columns.Contains(Nume_field) == false)
                                                {
                                                    Data_table_OD_attrib_existing.Columns.Add(Nume_field, typeof(String));

                                                }
                                            }

                                        }

                                        else
                                        {
                                            MessageBox.Show("Please reload your OD tables!");
                                            Freeze_operations = false;
                                            label_processing1.Visible = false;
                                            return;
                                        }

                                        System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                                        Nume_tables = Tables1.GetTableNames();

                                        int Id_od = 0;

                                        foreach (ObjectId Obj_ID1 in BTrecord)
                                        {
                                            Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                                            if (Ent1 != null)
                                            {
                                                Boolean Add_to_table = false;


                                                if (Ent1.Layer == Correct_layer)
                                                {
                                                    Add_to_table = true;
                                                }



                                                if (Add_to_table == true)
                                                {

                                                    bool Correct_table_exists = false;


                                                    Data_table_OD_attrib_existing.Rows.Add();


                                                    Data_table_OD_attrib_existing.Rows[Id_od]["OBJECT_ID"] = Ent1.ObjectId;
                                                    Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"] = 0;

                                                    string Type1 = Ent1.GetType().ToString();
                                                    Type1 = Type1.Replace("Autodesk.AutoCAD.DatabaseServices.", "");
                                                    Data_table_OD_attrib_existing.Rows[Id_od]["OBJECT_TYPE"] = Type1;

                                                    List_all_objId.Add(Ent1.ObjectId.ToString());



                                                    if (Tables1.IsTableDefined(Correct_table) == true)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.Table Tabla0 = Tables1[Correct_table];
                                                        using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla0.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                        {
                                                            if (Records1 != null)
                                                            {
                                                                if (Records1.Count > 0)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla0.FieldDefinitions;

                                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                                    {
                                                                        for (int i = 0; i < Record1.Count; ++i)
                                                                        {
                                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                            string Nume_field = Field_def1.Name;
                                                                            string Valoare1 = Record1[i].StrValue;
                                                                            Data_table_OD_attrib_existing.Rows[Id_od][Nume_field] = Valoare1;


                                                                        }
                                                                    }

                                                                    if (List_of_tables_on_layer.Contains(Correct_table) == false) List_of_tables_on_layer.Add(Correct_table);
                                                                    Correct_table_exists = true;
                                                                }

                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        MessageBox.Show("Please reload your OD tables!");
                                                        Freeze_operations = false;
                                                        return;
                                                    }

                                                    for (int k = 0; k < Nume_tables.Count; k = k + 1)
                                                    {
                                                        if (Tables1.IsTableDefined(Nume_tables[k]) == true)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.Table Tabla2 = Tables1[Nume_tables[k]];
                                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla2.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                                            {
                                                                if (Records1 != null)
                                                                {
                                                                    if (Records1.Count > 0)
                                                                    {
                                                                        Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"] = Convert.ToInt32(Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"]) + 1;
                                                                        if (List_of_tables_on_layer.Contains(Nume_tables[k]) == false)
                                                                        {
                                                                            List_of_tables_on_layer.Add(Nume_tables[k]);
                                                                        }
                                                                    }

                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (Convert.ToInt32(Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"]) > 1)
                                                    {

                                                        List_red.Add(Id_od);
                                                        List_red_objId.Add(Ent1.ObjectId.ToString());
                                                    }

                                                    if (Convert.ToInt32(Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"]) == 0)
                                                    {

                                                        List_yellow.Add(Id_od);
                                                        List_yellow_objId.Add(Ent1.ObjectId.ToString());
                                                    }
                                                    if (Convert.ToInt32(Data_table_OD_attrib_existing.Rows[Id_od]["OD_TABLE_COUNT"]) >= 1 & Correct_table_exists == false)
                                                    {

                                                        List_blue.Add(Id_od);
                                                        List_blue_objId.Add(Ent1.ObjectId.ToString());
                                                    }
                                                    Id_od = Id_od + 1;
                                                }
                                            }
                                        }

                                        Trans1.Commit();


                                    }
                                }
                                if (Data_table_OD_attrib_existing.Rows.Count > 0)
                                {

                                    string no_lines = "";



                                    DataGridView_data.DataSource = Data_table_OD_attrib_existing;


                                    DataGridView_data.AllowUserToAddRows = false;
                                    DataGridView_OD_data_Sorted(sender, e);
                                    DataGridView_data.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                                    no_lines = DataGridView_data.Rows.Count.ToString();

                                    if (List_red.Count > 0)
                                    {
                                        for (int k = 0; k < List_red.Count; ++k)
                                        {
                                            DataGridView_data.Rows[List_red[k]].DefaultCellStyle.BackColor = Color.Red;
                                        }
                                    }

                                    if (List_yellow.Count > 0)
                                    {
                                        for (int k = 0; k < List_yellow.Count; ++k)
                                        {
                                            DataGridView_data.Rows[List_yellow[k]].DefaultCellStyle.BackColor = Color.Yellow;
                                        }
                                    }

                                    if (List_blue.Count > 0)
                                    {
                                        for (int k = 0; k < List_blue.Count; ++k)
                                        {
                                            DataGridView_data.Rows[List_blue[k]].DefaultCellStyle.BackColor = Color.SkyBlue;
                                        }
                                    }



                                    textBox_no_rows.Text = no_lines;
                                    textBox_no_od_2.Text = List_red.Count.ToString();


                                    textBox_no_od_zero.Text = List_yellow.Count.ToString();

                                    textBox_no_wrong_od.Text = List_blue.Count.ToString();

                                    textBox_no_tables.Text = List_of_tables_on_layer.Count.ToString();




                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }



                    }
                }

                if (radioButton_BLOCKS.Checked == true)
                {


                    if (this.comboBox_layers_blocks_geomanager.Text == "")
                    {
                        MessageBox.Show("Please select attribute block!");
                        label_processing1.Visible = false;
                        return;
                    }


                    if (Update_data == true)
                    {

                        Freeze_operations = true;
                        try
                        {
                            DataGridView_data.DataSource = null;
                            DataGridView_data.Refresh();
                            textBox_no_rows.Text = "";
                            textBox_no_od_2.Text = "";
                            textBox_no_od_zero.Text = "";
                            textBox_no_tables.Text = "";
                            textBox_no_wrong_od.Text = "";

                            Data_table_OD_attrib_existing = new System.Data.DataTable();

                            {

                                List_all_objId = new List<string>();
                                List_update_objId = new List<ObjectId>();
                                List_update_row_index = new List<int>();

                                Data_table_OD_attrib_existing.Columns.Add("OBJECT_ID", typeof(ObjectId));
                                Data_table_OD_attrib_existing.Columns.Add("BLOCK_NAME", typeof(String));

                                label_processing1.Visible = true;
                                this.Refresh();



                                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                                {
                                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

                                        if (checkBox_user_selection.Checked == false)
                                        {
                                            foreach (ObjectId Obj_ID1 in BTrecord)
                                            {
                                                Entity Ent1 = (Entity)Trans1.GetObject(Obj_ID1, OpenMode.ForRead);
                                                if (Ent1 != null)
                                                {
                                                    Boolean Add_to_table = false;

                                                    BlockReference Block1 = null;

                                                    try
                                                    {
                                                        Block1 = (BlockReference)Ent1;
                                                    }
                                                    catch (System.Exception ex)
                                                    {

                                                    }



                                                    if (Block1 != null)
                                                    {
                                                        if (Functions.get_block_name(Block1) == comboBox_layers_blocks_geomanager.Text)
                                                        {
                                                            Add_to_table = true;
                                                        }
                                                    }



                                                    if (Add_to_table == true)
                                                    {
                                                        Data_table_OD_attrib_existing.Rows.Add();
                                                        Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1]["OBJECT_ID"] = Block1.ObjectId;
                                                        Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1]["BLOCK_NAME"] = comboBox_layers_blocks_geomanager.Text;


                                                        List_all_objId.Add(Block1.ObjectId.ToString());
                                                        System.Data.DataTable Table1 = Functions.Read_block_attributes_and_values(Block1);

                                                        for (int i = 0; i < Table1.Rows.Count; ++i)
                                                        {
                                                            string Atr_name = Table1.Rows[i]["ATTRIB"].ToString();
                                                            string Atr_value = Table1.Rows[i]["VALUE"].ToString();

                                                            if (Data_table_OD_attrib_existing.Columns.Contains(Atr_name) == false)
                                                            {
                                                                Data_table_OD_attrib_existing.Columns.Add(Atr_name);

                                                            }
                                                            Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1][Atr_name] = Atr_value;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Editor1.SetImpliedSelection(Empty_array);

                                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                            Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                            Prompt_rez.MessageForAdding = "\nSelect blocks:";
                                            Prompt_rez.SingleOnly = false;
                                            Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                                            if (Rezultat1.Status != PromptStatus.OK)
                                            {
                                                Freeze_operations = false;
                                                label_processing1.Visible = false;
                                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                                return;
                                            }


                                            for (int k = 0; k < Rezultat1.Value.Count; ++k)
                                            {
                                                Entity Ent1 = Trans1.GetObject(Rezultat1.Value[k].ObjectId, OpenMode.ForRead) as Entity;
                                                if (Ent1 != null)
                                                {
                                                    bool Add_to_table = false;

                                                    BlockReference Block1 = null;

                                                    try
                                                    {
                                                        Block1 = (BlockReference)Ent1;
                                                    }
                                                    catch (System.Exception)
                                                    {

                                                    }



                                                    if (Block1 != null)
                                                    {
                                                        if (Functions.get_block_name(Block1) == comboBox_layers_blocks_geomanager.Text)
                                                        {
                                                            Add_to_table = true;
                                                        }
                                                    }



                                                    if (Add_to_table == true)
                                                    {
                                                        Data_table_OD_attrib_existing.Rows.Add();
                                                        Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1]["OBJECT_ID"] = Block1.ObjectId;
                                                        Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1]["BLOCK_NAME"] = comboBox_layers_blocks_geomanager.Text;


                                                        List_all_objId.Add(Block1.ObjectId.ToString());
                                                        System.Data.DataTable Table1 = Functions.Read_block_attributes_and_values(Block1);

                                                        for (int i = 0; i < Table1.Rows.Count; ++i)
                                                        {
                                                            string Atr_name = Table1.Rows[i]["ATTRIB"].ToString();
                                                            string Atr_value = Table1.Rows[i]["VALUE"].ToString();

                                                            if (Data_table_OD_attrib_existing.Columns.Contains(Atr_name) == false)
                                                            {
                                                                Data_table_OD_attrib_existing.Columns.Add(Atr_name);

                                                            }
                                                            Data_table_OD_attrib_existing.Rows[Data_table_OD_attrib_existing.Rows.Count - 1][Atr_name] = Atr_value;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        Trans1.Commit();
                                    }
                                }

                                if (Data_table_OD_attrib_existing.Rows.Count > 0)
                                {

                                    DataGridView_data.DataSource = Data_table_OD_attrib_existing;

                                    DataGridView_data.Columns[0].ReadOnly = true;
                                    DataGridView_data.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
                                    DataGridView_data.Columns[1].ReadOnly = true;
                                    DataGridView_data.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;

                                    DataGridView_data.AllowUserToAddRows = false;
                                    DataGridView_OD_data_Sorted(sender, e);
                                    DataGridView_data.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                                    textBox_no_rows.Text = DataGridView_data.Rows.Count.ToString();



                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }



                    }



                }
                Freeze_operations = false;
                label_processing1.Visible = false;

            }



        }




        private void Button_Update_data_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
                {
                    MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                    label_processing1.Visible = false;
                    return;
                }


                Freeze_operations = true;
                try
                {


                    if (List_update_objId.Count > 0)
                    {
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                for (int i = 0; i < List_update_objId.Count; i = i + 1)
                                {

                                    //int Row_index = DataGridView_OD_data.CurrentCell.RowIndex;

                                    ObjectId Id1 = (ObjectId)List_update_objId[i];

                                    Entity Ent1 = null;
                                    try
                                    {
                                        Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);
                                    }
                                    catch (System.Exception ex)
                                    {

                                        MessageBox.Show("The object to be updated was deleted" + "\r\nRefresh!");
                                        Freeze_operations = false;
                                        return;
                                    }

                                    if (Ent1 != null)
                                    {
                                        if (radioButton_OD.Checked == true)
                                        {
                                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                            Autodesk.Gis.Map.ObjectData.Records Records1;
                                            if (Tables1.IsTableDefined(Correct_table) == true)
                                            {
                                                Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Correct_table];

                                                using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                {

                                                    if (Records1.Count > 0)
                                                    {
                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {

                                                            Autodesk.Gis.Map.Utilities.MapValue Valoare1;


                                                            for (int j = 4; j < Data_table_OD_attrib_existing.Columns.Count; ++j)
                                                            {

                                                                Valoare1 = Record1[j - 4];

                                                                if (Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j] != null)
                                                                {
                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Character)
                                                                    {
                                                                        Valoare1.Assign(Convert.ToString(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j]));
                                                                    }
                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Integer)
                                                                    {
                                                                        if (Functions.IsNumeric(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j].ToString()) == true)
                                                                        {
                                                                            Valoare1.Assign(Convert.ToInt32(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j]));
                                                                        }
                                                                    }
                                                                    if (Valoare1.Type == Autodesk.Gis.Map.Constants.DataType.Real)
                                                                    {
                                                                        if (Functions.IsNumeric(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j].ToString()) == true)
                                                                        {
                                                                            Valoare1.Assign(Convert.ToDouble(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j]));
                                                                        }
                                                                    }

                                                                }

                                                                Records1.UpdateRecord(Record1);
                                                            }

                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {

                                                MessageBox.Show("The table not found");
                                                Freeze_operations = false;
                                                return;
                                            }
                                        }

                                        if (radioButton_BLOCKS.Checked == true)
                                        {

                                            System.Collections.Specialized.StringCollection Col_name = new System.Collections.Specialized.StringCollection();
                                            System.Collections.Specialized.StringCollection Col_value = new System.Collections.Specialized.StringCollection();

                                            for (int j = 2; j < Data_table_OD_attrib_existing.Columns.Count; ++j)
                                            {
                                                if (Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j] != null)
                                                {
                                                    Col_name.Add(Data_table_OD_attrib_existing.Columns[j].ColumnName);
                                                    Col_value.Add(Data_table_OD_attrib_existing.Rows[List_update_row_index[i]][j].ToString());
                                                }

                                            }

                                            if (Col_name.Count > 0)
                                            {

                                                BlockReference Block1 = null;
                                                try
                                                {
                                                    Block1 = (BlockReference)Ent1;
                                                }
                                                catch (System.Exception EX)
                                                {

                                                }

                                                if (Block1 != null)
                                                {
                                                    Functions.Update_Attrib_block_values(Block1, Col_name, Col_value);

                                                }
                                            }

                                        }
                                    }
                                }
                                Trans1.Commit();
                                List_update_objId = new List<ObjectId>();
                                List_update_row_index = new List<int>();

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nothing to update");
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }
        }

        private void button_add_OD_table_and_remove_wrong_OD_Click(object sender, EventArgs e)
        {

            if (Data_table_OD_attrib_existing == null)
            {
                MessageBox.Show("No data loaded");
                return;
            }

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
            {
                MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                label_processing1.Visible = false;
                return;
            }

            int Error1 = 0;

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                try
                {
                    try
                    {
                        if (Data_table_OD_attrib_existing.Rows.Count > 0 & Correct_table != "")
                        {
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                    Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                    System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();

                                    Nume_tables = Tables1.GetTableNames();



                                    if (Nume_tables.Count == 0)
                                    {
                                        MessageBox.Show("Please refresh your data, object data table missing");
                                        DataGridView_data.DataSource = "";
                                        Data_table_OD_attrib_existing = new System.Data.DataTable();
                                        Freeze_operations = false;
                                        return;
                                    }


                                    if (Nume_tables.Count > 0)
                                    {
                                        Is_update_running = true;

                                        label_processing1.Visible = true;
                                        this.Refresh();

                                        for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; ++i)
                                        {
                                            ObjectId Id1 = (ObjectId)Data_table_OD_attrib_existing.Rows[i][0];
                                            if (List_blue_objId.Contains(Id1.ToString()) == true | List_yellow_objId.Contains(Id1.ToString()) == true | List_red_objId.Contains(Id1.ToString()) == true)
                                            {
                                                Autodesk.Gis.Map.ObjectData.Records Records1;
                                                Entity Ent1 = null;
                                                try
                                                {
                                                    Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    MessageBox.Show("Please refresh your data, objectID not existing");
                                                    DataGridView_data.DataSource = "";
                                                    Data_table_OD_attrib_existing = new System.Data.DataTable();
                                                    textBox_no_od_2.Text = "";
                                                    textBox_no_od_zero.Text = "";
                                                    textBox_no_rows.Text = "";
                                                    textBox_no_tables.Text = "";
                                                    textBox_no_wrong_od.Text = "";
                                                    Is_update_running = false;
                                                    Freeze_operations = false;
                                                    label_processing1.Visible = false;
                                                    return;
                                                }
                                                Error1 = 10;
                                                if (Ent1 != null)
                                                {
                                                    try
                                                    {

                                                        bool Exista_OD = false;

                                                        Data_table_OD_attrib_existing.Rows[i][1] = 1;

                                                        using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                        {
                                                            if (Records1 != null)
                                                            {
                                                                if (Records1.Count > 0)
                                                                // here I remove the object data for the case when the object has a different data table attached to it
                                                                {
                                                                    Error1 = 11;

                                                                    if (Tables1.IsTableDefined(Correct_table) == true)
                                                                    {
                                                                        Autodesk.Gis.Map.ObjectData.Table Tabla2 = Tables1[Correct_table];

                                                                        using (Autodesk.Gis.Map.ObjectData.Records Records2 = Tabla2.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                                        {
                                                                            if (Records2.Count > 0)
                                                                            {
                                                                                Exista_OD = true;
                                                                                Error1 = 1;
                                                                            }
                                                                        }

                                                                    }

                                                                    for (int k = 0; k < Nume_tables.Count; k = k + 1)
                                                                    {
                                                                        if (Nume_tables[k] != Correct_table)
                                                                        {
                                                                            Autodesk.Gis.Map.ObjectData.Table Tabla3 = Tables1[Nume_tables[k]];
                                                                            using (Autodesk.Gis.Map.ObjectData.Records Records3 = Tabla3.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                                            {
                                                                                if (Records3.Count > 0)
                                                                                {
                                                                                    System.Collections.IEnumerator ie = Records3.GetEnumerator();
                                                                                    while (ie.MoveNext())
                                                                                    {
                                                                                        Records3.RemoveRecord();
                                                                                        Error1 = 2;
                                                                                    }
                                                                                }
                                                                            }


                                                                        }
                                                                    }





                                                                }
                                                            }


                                                            if (Exista_OD == false)
                                                            {


                                                                Error1 = 3;

                                                                if (Tables1.IsTableDefined(Correct_table) == true)
                                                                {
                                                                    using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Correct_table])
                                                                    {
                                                                        using (Autodesk.Gis.Map.ObjectData.Record Record1 = Autodesk.Gis.Map.ObjectData.Record.Create())
                                                                        {
                                                                            Tabla1.InitRecord(Record1);
                                                                            Tabla1.AddRecord(Record1, Ent1.ObjectId);

                                                                            Error1 = 4;
                                                                        }
                                                                    }
                                                                }
                                                            }


                                                        }
                                                    }
                                                    catch (AccessViolationException ex1)
                                                    {
                                                        MessageBox.Show(ex1.Message);
                                                    }
                                                }


                                            }

                                        }

                                    }
                                    Trans1.Commit();
                                }

                            }
                            Error1 = 5;

                            for (int k = 0; k < DataGridView_data.Rows.Count; ++k)
                            {
                                DataGridView_data.Rows[k].DefaultCellStyle.BackColor = Color.White;
                                if (Table_filter != null)
                                {
                                    DataGridView_data[1, k].Value = 1;
                                }

                            }

                            if (DataGridView_data.DataSource == Table_filter)
                            {
                                DataGridView_data.DataSource = Data_table_OD_attrib_existing;
                            }

                            textBox_no_wrong_od.Text = "0";
                            textBox_no_tables.Text = "1";
                            textBox_no_od_zero.Text = "0";
                            textBox_no_od_2.Text = "0";
                            List_red = new List<int>();
                            List_yellow = new List<int>();
                            List_blue = new List<int>();
                            List_red_objId = new List<string>();
                            List_yellow_objId = new List<string>();
                            List_blue_objId = new List<string>();

                        }
                        else
                        {
                            MessageBox.Show("Please select the correct OD table");
                        }

                    }

                    catch (Autodesk.Gis.Map.MapException ex)
                    {
                        MessageBox.Show(ex.Message + vbcrlf + Error1.ToString());
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                label_processing1.Visible = false;
                Freeze_operations = false;
                Is_update_running = false;
            }

        }

        private void button_remove_od_Click(object sender, EventArgs e)
        {
            if (Data_table_OD_attrib_existing == null)
            {
                MessageBox.Show("No data loaded");
                return;
            }

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
            {
                MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                label_processing1.Visible = false;
                return;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_OD_attrib_existing.Rows.Count > 0)
                {
                    if (List_blue_objId.Count > 0 | List_yellow_objId.Count > 0 | List_red_objId.Count > 0)
                    {



                        for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; ++i)
                        {
                            ObjectId Id1 = (ObjectId)Data_table_OD_attrib_existing.Rows[i][0];
                            if (List_blue_objId.Contains(Id1.ToString()) == true | List_yellow_objId.Contains(Id1.ToString()) == true | List_red_objId.Contains(Id1.ToString()) == true)
                            {
                                try
                                {
                                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                                    {
                                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);



                                            Entity Ent1 = null;
                                            try
                                            {
                                                Ent1 = (Entity)Trans1.GetObject(Id1, OpenMode.ForWrite);
                                            }
                                            catch (System.Exception ex)
                                            {

                                            }

                                            if (Ent1 != null)
                                            {
                                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                                if (Tables1.IsTableDefined(Correct_table) == true)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Correct_table];


                                                    Autodesk.Gis.Map.ObjectData.Records Records1;


                                                    using (Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, true))
                                                    {
                                                        if (Records1.Count > 0)
                                                        {



                                                            System.Collections.IEnumerator ie = Records1.GetEnumerator();
                                                            while (ie.MoveNext())
                                                            {
                                                                Records1.RemoveRecord();
                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                            Trans1.Commit();







                                            textBox_no_rows.Text = Data_table_OD_attrib_existing.Rows.Count.ToString();



                                            textBox_no_od_2.Text = "0";


                                        }
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

                Freeze_operations = false;
            }
        }

        private void DataGridView_OD_data_Sorted(object sender, EventArgs e)
        {

            for (int j = 0; j < DataGridView_data.RowCount; j = j + 1)
            {
                String objid2 = DataGridView_data[0, j].Value.ToString();

                if (radioButton_OD.Checked == true)
                {
                    if (List_red.Count > 0)
                    {
                        if (List_red_objId.Contains(objid2) == true)
                        {
                            DataGridView_data.Rows[j].DefaultCellStyle.BackColor = Color.Red;
                        }
                    }

                    if (List_yellow.Count > 0)
                    {
                        if (List_yellow_objId.Contains(objid2) == true)
                        {
                            DataGridView_data.Rows[j].DefaultCellStyle.BackColor = Color.Yellow;
                        }
                    }


                    if (List_blue.Count > 0)
                    {
                        if (List_blue_objId.Contains(objid2) == true)
                        {
                            DataGridView_data.Rows[j].DefaultCellStyle.BackColor = Color.SkyBlue;
                        }
                    }
                }

                if (Object_id_grid_current == objid2)
                {
                    DataGridView_data.CurrentCell = DataGridView_data[0, j];

                }
            }
        }

        private void DataGrid_od_data_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {
                    string Id2 = DataGridView_data[0, DataGridView_data.CurrentCell.RowIndex].Value.ToString();

                    if (Id2 != Object_id_grid_current)
                    {
                        Object_id_grid_current = Id2;

                    }
                }
                catch (System.StackOverflowException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }
        }

        private void button_go_to_table_row_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
                {
                    MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                    label_processing1.Visible = false;
                    return;
                }
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                            bool ask_for_selection = false;
                            Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_object = (Autodesk.AutoCAD.EditorInput.PromptSelectionResult)Editor1.SelectImplied();

                            if (Rezultat_object.Status == PromptStatus.OK)
                            {
                                if (Rezultat_object.Value.Count == 0)
                                {
                                    ask_for_selection = true;
                                }
                                if (Rezultat_object.Value.Count > 1)
                                {
                                    MessageBox.Show("There is more than one object selected," + "\r\n" + "the first object in selection will be the one that will be current in table");
                                    ask_for_selection = false;
                                }
                            }
                            else ask_for_selection = true;



                            if (ask_for_selection == true)
                            {
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_object = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_object.MessageForAdding = "\nSelect an object";
                                Prompt_object.SingleOnly = true;
                                Rezultat_object = Editor1.GetSelection(Prompt_object);
                            }


                            if (Rezultat_object.Status != PromptStatus.OK)
                            {
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                Freeze_operations = false;
                                return;
                            }

                            Entity Ent1 = (Entity)Trans1.GetObject(Rezultat_object.Value[0].ObjectId, OpenMode.ForRead);
                            ObjectId Id1 = Ent1.ObjectId;

                            if (Data_table_OD_attrib_existing != null)
                            {
                                if (Data_table_OD_attrib_existing.Rows.Count > 0)
                                {
                                    for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; i = i + 1)
                                    {
                                        if (Data_table_OD_attrib_existing.Rows[i]["OBJECT_ID"] != null)
                                        {
                                            ObjectId iD2 = (ObjectId)Data_table_OD_attrib_existing.Rows[i]["OBJECT_ID"];
                                            if (Id1 == iD2)
                                            {

                                                for (int j = 0; j < DataGridView_data.Rows.Count; j = j + 1)
                                                {
                                                    if (iD2.ToString() == DataGridView_data[0, j].Value.ToString())
                                                    {
                                                        DataGridView_data.CurrentCell = DataGridView_data[0, j];
                                                        DataGridView_data.Refresh();
                                                        j = DataGridView_data.Rows.Count;
                                                        i = Data_table_OD_attrib_existing.Rows.Count;

                                                    }
                                                }



                                            }

                                        }

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
            }
        }

        private void button_zoom_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
                {
                    MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                    label_processing1.Visible = false;
                    return;
                }
                Freeze_operations = true;
                try
                {
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                            int Row_index_data_table = 0;
                            int R_i = DataGridView_data.CurrentCell.RowIndex;
                            string Object_id_grid = DataGridView_data[0, R_i].Value.ToString();

                            for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; i = i + 1)
                            {

                                if (Data_table_OD_attrib_existing.Rows[i][0] != DBNull.Value)
                                {
                                    string Object_id = Data_table_OD_attrib_existing.Rows[i][0].ToString();
                                    if (Object_id == Object_id_grid)
                                    {
                                        Row_index_data_table = i;
                                        i = Data_table_OD_attrib_existing.Rows.Count;
                                    }
                                }
                            }

                            ObjectId ObjId = (ObjectId)Data_table_OD_attrib_existing.Rows[Row_index_data_table][0];
                            try
                            {
                                Entity Ent1 = (Entity)Trans1.GetObject(ObjId, OpenMode.ForRead);
                                if (Ent1 != null)
                                {
                                    Point3d minx = new Point3d();
                                    Point3d maxx = new Point3d();
                                    try
                                    {
                                        minx = Ent1.GeometricExtents.MinPoint;
                                        maxx = Ent1.GeometricExtents.MaxPoint;
                                    }
                                    catch (System.Exception)
                                    {
                                        if (Ent1 is BlockReference)
                                        {
                                            BlockReference bl1 = Ent1 as BlockReference;
                                            minx = bl1.GeometryExtentsBestFit().MinPoint;
                                            maxx = bl1.GeometryExtentsBestFit().MaxPoint;
                                        }
                                    }


                                    using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                                    {

                                        int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                        //from here 2015 dlls:
                                        Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                        kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                        Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                        // to here 2015 dlls

                                        //from here 2013 dlls:

                                        //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                        // to here 2013 dlls

                                        if (view != null)
                                        {
                                            using (view)
                                            {
                                                if (Ent1 is DBPoint)
                                                {

                                                    //if (List_circles_zoom_objId.Count > 0)
                                                    //{
                                                    //foreach (ObjectId m in List_circles_zoom_objId)
                                                    //{

                                                    //Entity C2 = (Entity)Trans1.GetObject(m, OpenMode.ForWrite);
                                                    // C2.Erase();
                                                    //}
                                                    //}

                                                    //List_circles_zoom_objId = new List<ObjectId>();

                                                    //TypedValue[] typedV1 = new TypedValue[4];
                                                    //typedV1.SetValue(new TypedValue((int)DxfCode.Operator, "<and"), 0);
                                                    //typedV1.SetValue(new TypedValue((int)DxfCode.LayerName, "Zoom_circles"), 1);
                                                    //typedV1.SetValue(new TypedValue((int)DxfCode.Start, "CIRCLE"),2);
                                                    //typedV1.SetValue(new TypedValue((int)DxfCode.Operator, "and>"), 3);
                                                    //SelectionFilter sf = new SelectionFilter(typedV1);
                                                    //PromptSelectionResult acSSPrompt;
                                                    //acSSPrompt = Editor1.SelectAll(sf);
                                                    //if (acSSPrompt.Status == PromptStatus.OK)
                                                    //{
                                                    //SelectionSet acSSet = acSSPrompt.Value;
                                                    //if (acSSet.Count > 0)
                                                    //{

                                                    //}

                                                    //}

                                                    Functions.Creaza_layer("Zoom_circles", 40, false);
                                                    DBPoint Pt1 = new DBPoint();

                                                    try
                                                    {
                                                        Pt1 = (DBPoint)Ent1;
                                                    }
                                                    catch (System.Exception EX)
                                                    {
                                                    }


                                                    Circle C1 = new Circle(Pt1.Position, Vector3d.ZAxis, 200);
                                                    C1.Layer = "Zoom_circles";
                                                    BTrecord.AppendEntity(C1);
                                                    Trans1.AddNewlyCreatedDBObject(C1, true);

                                                    view.ZoomExtents(C1.GeometricExtents.MaxPoint, C1.GeometricExtents.MinPoint);
                                                }
                                                else
                                                {
                                                    view.ZoomExtents(maxx, minx);
                                                }
                                                view.Zoom(0.95);//<--optional 
                                                GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                            }
                                        }
                                        Trans1.Commit();
                                    }
                                    Ent1.Highlight();
                                    ObjectId[] objid1;
                                    objid1 = new ObjectId[1];
                                    objid1[0] = Ent1.ObjectId;
                                    Editor1.SetImpliedSelection(objid1);
                                }
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }
                    }
                }







                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Freeze_operations = false;
            }

        }






        private void button_export_to_excel_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
            {
                MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                label_processing1.Visible = false;
                return;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                try
                {

                    if (DataGridView_data.Rows.Count > 0)
                    {
                        label_processing1.Visible = true;
                        this.Refresh();

                        System.Data.DataTable dt = new System.Data.DataTable();


                        foreach (DataGridViewColumn column in DataGridView_data.Columns)
                        {

                            dt.Columns.Add(column.HeaderText, typeof(string));

                        }

                        dt.Rows.Add();

                        foreach (DataGridViewColumn column in DataGridView_data.Columns)
                        {

                            dt.Rows[0][column.Index] = column.HeaderText.ToString();

                        }

                        foreach (DataGridViewRow row in DataGridView_data.Rows)
                        {



                            dt.Rows.Add();

                            foreach (DataGridViewCell cell in row.Cells)
                            {

                                string valoare1 = cell.Value.ToString();
                                if (radioButton_BLOCKS.Checked == true) valoare1 = valoare1.Replace("\\P", " ");
                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = valoare1;

                            }

                        }

                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_NEW_worksheet_from_Excel();


                        W1.Cells.NumberFormat = "@";


                        int maxRows = dt.Rows.Count, maxCols = dt.Columns.Count;
                        Microsoft.Office.Interop.Excel.Range range = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                        object[,] values = new object[maxRows, maxCols];
                        for (int row = 0; row < maxRows; row++)
                        {
                            for (int col = 0; col < maxCols; col++)
                            {
                                if (dt.Rows[row][col] != DBNull.Value)
                                {
                                    values[row, col] = dt.Rows[row][col];
                                }
                            }
                        }
                        range.Value2 = values;
                    }
                }

                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                label_processing1.Visible = false;
                Freeze_operations = false;
            }
        }

        private void button_import_from_excel_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
            {
                MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                label_processing1.Visible = false;
                return;
            }
            if (Freeze_operations == false)
            {
                System.Diagnostics.Process[] process1 = System.Diagnostics.Process.GetProcessesByName("EXCEL");

                if (process1.Length == 0)
                {
                    MessageBox.Show("No open Excel spreadsheet found." + vbcrlf + "Import Cancelled");
                    return;
                }

                if (process1.Length > 1)
                {
                    MessageBox.Show("More than one Excel spreadsheet opened." + vbcrlf + "Import Cancelled");
                    return;
                }

                if (process1.Length == 1)
                {
                    if (Functions.Get_no_of_workbooks_from_Excel() > 1)
                    {
                        MessageBox.Show("More than one Excel spreadsheet opened." + vbcrlf + "Import Cancelled");
                        return;
                    }
                }


                if (Data_table_OD_attrib_existing == null)
                {
                    if (Data_table_OD_attrib_existing.Rows.Count == 0)
                    {
                        MessageBox.Show("No data has been loaded from the current drawing." + vbcrlf + "Import Cancelled");
                        return;
                    }
                }



                Freeze_operations = true;
                label_processing1.Visible = true;
                this.Refresh();

                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_active_worksheet_from_Excel();

                    int start1 = 1;
                    int end1 = 100000;

                    int End_Col1 = W1.Range["ZZ1"].Column;

                    System.Data.DataTable dt = new System.Data.DataTable();
                    for (int j = 1; j <= End_Col1; j = j + 1)
                    {
                        string Valoare = W1.Cells[1, j].Value2;
                        if (Valoare == null)
                        {

                            End_Col1 = j - 1;
                        }
                        else
                        {
                            dt.Columns.Add(Valoare, typeof(string));
                        }
                    }

                    if (End_Col1 == 0)
                    {
                        MessageBox.Show("Cell A1 is empty!" + vbcrlf + "Import Cancelled");
                        label_processing1.Visible = false;
                        Freeze_operations = false;
                        return;
                    }

                    for (int i = start1 + 1; i <= end1; i = i + 1)
                    {
                        if (W1.Range["A" + i.ToString()].Value2 == null)
                        {
                            end1 = i - 1;
                        }
                        else
                        {
                            dt.Rows.Add();
                        }
                    }

                    if (end1 == 0)
                    {
                        MessageBox.Show("Cell A1 is empty!" + vbcrlf + "Import Cancelled");
                        label_processing1.Visible = false;
                        Freeze_operations = false;
                        return;
                    }

                    Microsoft.Office.Interop.Excel.Range range = W1.Range[W1.Cells[2, 1], W1.Cells[end1, End_Col1]];
                    object[,] values = new object[end1 - 1, End_Col1];

                    values = range.Value2;

                    for (int i = 0; i < dt.Rows.Count; ++i)
                    {
                        for (int j = 0; j < dt.Columns.Count; ++j)
                        {
                            dt.Rows[i][j] = values[i + 1, j + 1];
                        }
                    }

                    bool New_list_update_reset = true;
                    int Rows_updated = 0;

                    string Object_id_column = "OBJECT_ID";
                    for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; ++i)
                    {
                        string Obj_id1 = Data_table_OD_attrib_existing.Rows[i][Object_id_column].ToString();
                        if (dt.Columns.Contains(Object_id_column) == true)
                        {
                            for (int j = 0; j < dt.Rows.Count; ++j)
                            {
                                string Obj_id2 = dt.Rows[j][Object_id_column].ToString();

                                if (Obj_id1 == Obj_id2)
                                {
                                    for (int m = 1; m < Data_table_OD_attrib_existing.Columns.Count; ++m)
                                    {
                                        string Column1 = Data_table_OD_attrib_existing.Columns[m].ColumnName.ToString();
                                        for (int n = 1; n < dt.Columns.Count; ++n)
                                        {
                                            string Column2 = dt.Columns[n].ColumnName.ToString();
                                            if (Column1 == Column2)
                                            {

                                                string Val1 = Data_table_OD_attrib_existing.Rows[i][m].ToString();
                                                string Val2 = dt.Rows[j][n].ToString();
                                                if (Val1 != Val2)
                                                {
                                                    if (New_list_update_reset == true)
                                                    {
                                                        List_update_row_index = new List<int>();
                                                        List_update_objId = new List<ObjectId>();
                                                        New_list_update_reset = false;
                                                    }



                                                    if (List_update_row_index.Contains(i) == false)
                                                    {
                                                        List_update_row_index.Add(i);
                                                        List_update_objId.Add((ObjectId)Data_table_OD_attrib_existing.Rows[i][Object_id_column]);
                                                        Rows_updated = Rows_updated + 1;
                                                    }
                                                    Data_table_OD_attrib_existing.Rows[i][m] = Val2;

                                                }


                                            }

                                        }
                                    }




                                    j = dt.Rows.Count;
                                }

                            }
                        }
                        else
                        {
                            MessageBox.Show("Excel does not contain OBJECT_ID Column!" + vbcrlf + "Import Cancelled");
                            label_processing1.Visible = false;
                            Freeze_operations = false;
                            return;
                        }
                    }


                    if (DataGridView_data.DataSource != Data_table_OD_attrib_existing)
                    {

                        if (Data_table_OD_attrib_existing.Rows.Count > 0)
                        {

                            string no_lines = "";



                            DataGridView_data.DataSource = Data_table_OD_attrib_existing;

                            DataGridView_data.Columns[0].ReadOnly = true;
                            DataGridView_data.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
                            DataGridView_data.Columns[1].ReadOnly = true;
                            DataGridView_data.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
                            DataGridView_data.Columns[2].ReadOnly = true;
                            DataGridView_data.Columns[2].DefaultCellStyle.BackColor = Color.LightGray;
                            DataGridView_data.Columns[3].ReadOnly = true;
                            DataGridView_data.Columns[3].DefaultCellStyle.BackColor = Color.LightGray;


                            for (int k = 4; k < DataGridView_data.ColumnCount; ++k)
                            {
                                string Column_name = DataGridView_data.Columns[k].Name;
                                if (Column_name.ToUpper() == "FEATID" | Column_name.ToUpper() == "MMID" | Column_name.ToUpper() == "HMMID")
                                {
                                    DataGridView_data.Columns[k].ReadOnly = true;
                                    DataGridView_data.Columns[k].DefaultCellStyle.BackColor = Color.LightGray;
                                }
                            }

                            DataGridView_data.AllowUserToAddRows = false;
                            DataGridView_OD_data_Sorted(sender, e);
                            DataGridView_data.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                            no_lines = DataGridView_data.Rows.Count.ToString();

                            if (List_red.Count > 0)
                            {
                                for (int k = 0; k < List_red.Count; ++k)
                                {
                                    DataGridView_data.Rows[List_red[k]].DefaultCellStyle.BackColor = Color.Red;
                                }
                            }

                            if (List_yellow.Count > 0)
                            {
                                for (int k = 0; k < List_yellow.Count; ++k)
                                {
                                    DataGridView_data.Rows[List_yellow[k]].DefaultCellStyle.BackColor = Color.Yellow;
                                }
                            }

                            if (List_blue.Count > 0)
                            {
                                for (int k = 0; k < List_blue.Count; ++k)
                                {
                                    DataGridView_data.Rows[List_blue[k]].DefaultCellStyle.BackColor = Color.SkyBlue;
                                }
                            }



                            textBox_no_rows.Text = no_lines;
                            textBox_no_od_2.Text = List_red.Count.ToString();


                            textBox_no_od_zero.Text = List_yellow.Count.ToString();

                            textBox_no_wrong_od.Text = List_blue.Count.ToString();

                            textBox_no_tables.Text = List_of_tables_on_layer.Count.ToString();




                        }
                    }
                    string Suffix = "s";
                    if (Rows_updated == 1) Suffix = "";
                    MessageBox.Show(Rows_updated.ToString() + " row" + Suffix + " updated");


                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                label_processing1.Visible = false;
                Freeze_operations = false;
            }
        }

        private void button_Filter_Click(object sender, EventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

            if (ThisDrawing.Database.OriginalFileName != label_drawing_name.Text)
            {
                MessageBox.Show("You press the Load Data button into another drawing!" + vbcrlf + "No operation executed");
                label_processing1.Visible = false;
                return;
            }

            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                if (DataGridView_data.RowCount > 0)
                {
                    if (List_blue_objId.Count > 0 | List_yellow_objId.Count > 0 | List_red_objId.Count > 0)
                    {

                        Table_filter = new System.Data.DataTable();
                        Table_filter.Columns.Add("OBJECT_ID", typeof(ObjectId));
                        Table_filter.Columns.Add("OD_TABLE_COUNT", typeof(int));
                        Table_filter.Columns.Add("OBJECT_TYPE", typeof(String));
                        Table_filter.Columns.Add("BLOCK_NAME", typeof(String));

                        for (int k = 4; k < Data_table_OD_attrib_existing.Columns.Count; ++k)
                        {
                            Table_filter.Columns.Add(Data_table_OD_attrib_existing.Columns[k].ColumnName, typeof(String));
                        }

                        int j = 0;

                        for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; ++i)
                        {
                            string Id1 = Data_table_OD_attrib_existing.Rows[i][0].ToString();
                            if (List_blue_objId.Contains(Id1) == true | List_yellow_objId.Contains(Id1) == true | List_red_objId.Contains(Id1) == true)
                            {
                                Table_filter.Rows.Add();

                                for (int k = 0; k < Data_table_OD_attrib_existing.Columns.Count; ++k)
                                {
                                    if (Data_table_OD_attrib_existing.Rows[i][k] != DBNull.Value)
                                    {
                                        Table_filter.Rows[j][k] = Data_table_OD_attrib_existing.Rows[i][k];
                                    }
                                }

                                j = j + 1;

                            }

                        }


                        textBox_no_rows.Text = DataGridView_data.RowCount.ToString();
                        DataGridView_data.DataSource = Table_filter;


                        for (int i = 0; i < Table_filter.Rows.Count; ++i)
                        {
                            string Id1 = Table_filter.Rows[i][0].ToString();
                            if (List_red_objId.Contains(Id1) == true)
                            {
                                DataGridView_data.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                            }
                            if (List_yellow_objId.Contains(Id1) == true)
                            {
                                DataGridView_data.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                            }
                            if (List_blue_objId.Contains(Id1) == true)
                            {
                                DataGridView_data.Rows[i].DefaultCellStyle.BackColor = Color.SkyBlue;
                            }

                        }

                        DataGridView_data.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("There are no issues in the layer");
                    }

                }
                Freeze_operations = false;
            }
        }

        private void DataGridView_data_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            if (Is_update_running == false)
            {

                int R_i = DataGridView_data.CurrentCell.RowIndex;

                if (DataGridView_data.Rows[R_i].DefaultCellStyle.BackColor == Color.Red | DataGridView_data.Rows[R_i].DefaultCellStyle.BackColor == Color.Yellow | DataGridView_data.Rows[R_i].DefaultCellStyle.BackColor == Color.SkyBlue)
                {
                    DataGridView_data.CurrentCell.Value = "";
                }
                else
                {
                    string ID1 = DataGridView_data[0, R_i].Value.ToString();

                    if (List_all_objId.Contains(ID1) == true)
                    {
                        int T_i = List_all_objId.IndexOf(ID1);

                        if (List_update_row_index.Contains(T_i) == false)
                        {

                            List_update_objId.Add((ObjectId)Data_table_OD_attrib_existing.Rows[T_i][0]);
                            List_update_row_index.Add(T_i);
                        }

                        string Updated_val = DataGridView_data[DataGridView_data.CurrentCell.ColumnIndex, R_i].Value.ToString();
                        Data_table_OD_attrib_existing.Rows[T_i][DataGridView_data.CurrentCell.ColumnIndex] = Updated_val;
                    }
                }
            }
        }

        private void Button_MouseEnter(object sender, EventArgs e)
        {
            if (sender is System.Windows.Forms.Button)
            {
                System.Windows.Forms.Button Button1 = (System.Windows.Forms.Button)sender;
                Button1.UseVisualStyleBackColor = false;
                Button1.BackColor = Color.DodgerBlue;
            }
        }

        private void Button_MouseLeave(object sender, EventArgs e)
        {
            if (sender is System.Windows.Forms.Button)
            {
                System.Windows.Forms.Button Button1 = (System.Windows.Forms.Button)sender;
                Button1.UseVisualStyleBackColor = true;
                Button1.BackColor = Color.DimGray;
            }
        }

        private void panel_logo_DoubleClick(object sender, EventArgs e)
        {
            if (panel_excel.Visible == true)
            {
                panel_excel.Visible = false;
                panel_user_select.Visible = false;
            }
            else
            {
                panel_excel.Visible = true;
                panel_user_select.Visible = true;
            }
        }

        /// BACKUPS

        private void button_export_to_excel_Click_backup(object sender, EventArgs e)
        {
            try
            {
                if (Data_table_OD_attrib_existing != null)
                {
                    if (Data_table_OD_attrib_existing.Rows.Count > 0)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_NEW_worksheet_from_Excel();

                        for (int i = 0; i < Data_table_OD_attrib_existing.Columns.Count; i = i + 1)
                        {
                            W1.Cells[1, i + 1].NumberFormat = "@";
                            W1.Cells[1, i + 1].value2 = Data_table_OD_attrib_existing.Columns[i].ColumnName;
                        }

                        int No_of_cells = DataGridView_data.GetCellCount(DataGridViewElementStates.Selected);



                        if (No_of_cells <= 1)
                        {
                            for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; i = i + 1)
                            {

                                for (int j = 0; j < Data_table_OD_attrib_existing.Columns.Count; j = j + 1)
                                {
                                    if (Data_table_OD_attrib_existing.Rows[i][j] != DBNull.Value)
                                    {
                                        W1.Cells[i + 2, j + 1].NumberFormat = "@";
                                        W1.Cells[i + 2, j + 1].Value2 = Convert.ToString(Data_table_OD_attrib_existing.Rows[i][j]);
                                    }
                                }
                            }
                        }
                        else
                        {
                            List<string> List1 = new List<string>();
                            for (int i = 0; i <= No_of_cells - 1; i = i + 1)
                            {
                                int R_i = DataGridView_data.SelectedCells[i].RowIndex;
                                string Object_id = DataGridView_data[0, R_i].Value.ToString();

                                if (List1.Contains(Object_id) == false)
                                {
                                    List1.Add(Object_id);
                                }
                            }

                            int Index_row_excel = 2;

                            for (int i = 0; i < Data_table_OD_attrib_existing.Rows.Count; i = i + 1)
                            {

                                if (Data_table_OD_attrib_existing.Rows[i][0] != DBNull.Value)
                                {
                                    string Object_id = Data_table_OD_attrib_existing.Rows[i][0].ToString();
                                    if (List1.Contains(Object_id) == true)
                                    {
                                        for (int j = 0; j < Data_table_OD_attrib_existing.Columns.Count; j = j + 1)
                                        {
                                            if (Data_table_OD_attrib_existing.Rows[i][j] != DBNull.Value)
                                            {
                                                W1.Cells[Index_row_excel, j + 1].NumberFormat = "@";
                                                W1.Cells[Index_row_excel, j + 1].Value2 = Convert.ToString(Data_table_OD_attrib_existing.Rows[i][j]);
                                            }
                                        }
                                        Index_row_excel = Index_row_excel + 1;
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
        }

        private void button_import_from_excel_Click_backup(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_active_worksheet_from_Excel();
                String Column_start = "A";
                String Column_end = "ZZ";
                int start1 = 1;
                int end1 = 100000;



                int Col1 = W1.Range[Column_start + Convert.ToString(1)].Column;
                int Col2 = W1.Range[Column_end + Convert.ToString(1)].Column;
                for (int j = Col1 + 1; j <= Col2; j = j + 1)
                {
                    string Valoare = W1.Cells[1, j].Value2;
                    if (Valoare == null)
                    {

                        Col2 = j - 1;
                    }
                }

                for (int i = start1 + 1; i <= end1; i = i + 1)
                {
                    if (W1.Range["A" + i.ToString()].Value2 == null)
                    {
                        end1 = i - 1;
                    }
                }

                string Object_id_column = "OBJECT_ID";

                for (int i = start1 + 1; i <= end1; i = i + 1)
                {
                    String Object_id_excel = Convert.ToString(W1.Range[Column_start + i.ToString()].Value2);
                    for (int k = 0; k < Data_table_OD_attrib_existing.Rows.Count; k = k + 1)
                    {
                        String Object_id_OD_table = Convert.ToString(Data_table_OD_attrib_existing.Rows[k][Object_id_column]);
                        if (Object_id_excel == Object_id_OD_table)
                        {
                            for (int j = Col1 + 1; j <= Col2; j = j + 1)
                            {
                                String Excel_column = Convert.ToString(W1.Cells[start1, j].Value2);
                                if (Data_table_OD_attrib_existing.Columns.Contains(Excel_column) == true)
                                {
                                    Data_table_OD_attrib_existing.Rows[k][Excel_column] = Convert.ToString(W1.Cells[i, j].Value2);
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


        private void radioButton_OD_blocks_CheckedChanged(object sender, EventArgs e)
        {


            if (radioButton_OD.Checked == true)
            {
                button_Filter.Visible = true;
                button_add_OD_table.Visible = true;
                comboBox_od_existing_tables.Visible = true;
                label_correct_od_table.Visible = true;
                textBox_OD_TABLES.Text = "Total Number of OD Tables on Layer:";
                textBox_INCORRECT_od.Text = "Features with Incorrect OD Tables:";
                textBox_MultipleOD.Text = "Features with Multiple OD Tables:";
                textBox_missing_OD.Text = "Features with Missing OD Tables:";
                textBox_Features.Text = "Features: ";
                textBox_no_od_zero.Text = "";
                textBox_no_od_2.Text = "";
                textBox_no_wrong_od.Text = "";
                textBox_no_tables.Text = "";
                textBox_no_rows.Text = "";
                button_zoom_row_object_data.Text = "Select Feature";
                button_zoom.Text = "Zoom To Feature";
                label_current_layer_block.Text = "Current Layer";
                label_od_block_table.Text = "Object Data Table";
                label_Apply_Changes.Text = "Edits Will Only Be Pushed to Features When Apply Changes Has Been Selected.";
            }

            if (radioButton_BLOCKS.Checked == true)
            {
                button_Filter.Visible = false;
                button_add_OD_table.Visible = false;
                comboBox_od_existing_tables.Visible = false;
                label_correct_od_table.Visible = false;
                textBox_OD_TABLES.Text = "";
                textBox_INCORRECT_od.Text = "";
                textBox_MultipleOD.Text = "";
                textBox_missing_OD.Text = "";
                textBox_Features.Text = "Blocks: ";
                textBox_no_od_zero.Text = "";
                textBox_no_od_2.Text = "";
                textBox_no_wrong_od.Text = "";
                textBox_no_tables.Text = "";
                textBox_no_rows.Text = "";
                button_zoom_row_object_data.Text = "Select Block";
                button_zoom.Text = "Zoom To Block";
                label_current_layer_block.Text = "Current Block";
                label_od_block_table.Text = "Block Attributes Table";
                label_Apply_Changes.Text = "Edits Will Only Be Pushed to Blocks When Apply Changes Has Been Selected.";
            }
            DataGridView_data.DataSource = "";
            DataGridView_data.Refresh();
            comboBox_od_existing_tables.Items.Clear();
            comboBox_layers_blocks_geomanager.Items.Clear();

            if (radioButton_OD.Checked == true)
            {
                comboBox_layers_blocks_geomanager.Width = Oldw1;
                comboBox_od_existing_tables.Left = Oldl2;
                label_correct_od_table.Left = Oldll2;
                button_refresh_grid.Left = Oldl3;
            }

            if (radioButton_BLOCKS.Checked == true)
            {

                comboBox_layers_blocks_geomanager.Width = Oldw1;

                button_refresh_grid.Left = Oldl2;
            }

            this.Refresh();


        }




        /// OVERLAP FORM CODE

        private void Button_add_new_combobox_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {

                    System.Windows.Forms.ComboBox Combo_1 = new System.Windows.Forms.ComboBox();
                    Combo_1.Left = comboBox_GIC_Layer1.Left;
                    Combo_1.Size = comboBox_GIC_Layer1.Size;
                    Combo_1.BackColor = comboBox_GIC_Layer1.BackColor;
                    Combo_1.DropDownStyle = comboBox_GIC_Layer1.DropDownStyle;
                    Combo_1.FlatStyle = comboBox_GIC_Layer1.FlatStyle;
                    Combo_1.ForeColor = comboBox_GIC_Layer1.ForeColor;
                    Combo_1.FormattingEnabled = comboBox_GIC_Layer1.FormattingEnabled;

                    Combo_1.Top = Convert.ToInt32(Ultimul_top + Spacing);

                    panel_layerpanel.Controls.Add(Combo_1);

                    Functions.Incarca_existing_layers_to_combobox(Combo_1);
                    Ultimul_top = Combo_1.Top;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
                Freeze_operations = false;
            }

        }

        private void Button_remove_combo_Click(object sender, EventArgs e)
        {


            if (comboBox_GIC_Layer1.Top <= Ultimul_top - Spacing)
            {

                if (Freeze_operations == false)
                {
                    Freeze_operations = true;

                    System.Windows.Forms.Control[] Del1;
                    int Index1 = 0;

                    Del1 = new System.Windows.Forms.Control[Index1 + 1];

                    foreach (System.Windows.Forms.Control control1 in panel_layerpanel.Controls)
                        if (control1 is System.Windows.Forms.ComboBox)
                        {
                            {

                                //MessageBox.Show((control1.Top).ToString());
                                if (control1.Top > Ultimul_top - Spacing)
                                {
                                    Array.Resize(ref Del1, Index1 + 1);
                                    Del1[Index1] = control1;
                                    Index1 = Index1 + 1;
                                }
                                else
                                {
                                    System.Windows.Forms.ComboBox Combo1 = (System.Windows.Forms.ComboBox)control1;
                                }
                            }
                        }


                    if (Del1 != null)
                    {
                        for (int i = 0; i < Del1.Length; ++i)
                        {
                            panel_layerpanel.Controls.Remove(Del1[i]);

                        }
                    }

                    if (comboBox_GIC_Layer1.Top <= Ultimul_top)
                    {
                        Ultimul_top = Ultimul_top - Spacing;
                    }
                    this.Refresh();
                    Freeze_operations = false;
                }
            }
        }

        private void Button_done_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                String No_plot = "NO PLOT";



                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;





                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;


                Editor1.SetImpliedSelection(Empty_array);
                try
                {
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();





                    using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                    {

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);

                            System.Collections.Specialized.StringCollection Lista1 = new System.Collections.Specialized.StringCollection();

                            foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                            {
                                if (Combo1 is System.Windows.Forms.ComboBox)
                                {
                                    String Text1 = Combo1.Text;
                                    if (Text1 != "")
                                    {
                                        if (Lista1.Contains(Text1) == false)
                                        {
                                            Lista1.Add(Text1);
                                        }
                                    }
                                }
                            }
                            if (Lista1.Count > 0)
                            {
                                foreach (ObjectId objID in BTrecord)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(objID, OpenMode.ForRead);

                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.RotatedDimension & Ent1.Layer == No_plot)
                                    {
                                        Ent1.UpgradeOpen();
                                        Ent1.Erase();
                                    }

                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Region & Ent1.Layer == No_plot)
                                    {
                                        Ent1.UpgradeOpen();
                                        Ent1.Erase();
                                    }
                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                    {
                                        Polyline Poly1 = (Polyline)Ent1;
                                        if (Lista1.Contains(Poly1.Layer) == true)
                                        {
                                            Poly1.UpgradeOpen();
                                            Poly1.ColorIndex = 256;
                                        }
                                    }

                                }
                            }



                            Trans1.Commit();


                        }
                    }






                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
                Editor1.WriteMessage("\nCommand:");

            }
        }

        private void button_GIC_Refresh_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                try
                {

                    foreach (System.Windows.Forms.Control Control1 in panel_layerpanel.Controls)
                    {
                        if (Control1 is System.Windows.Forms.ComboBox)
                        {
                            System.Windows.Forms.ComboBox Combo_1 = (System.Windows.Forms.ComboBox)Control1;
                            Functions.Incarca_existing_layers_to_combobox(Combo_1);
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
                Freeze_operations = false;
            }
        }

        private void button_GIC_Detect_Open_PL_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

                try
                {
                    using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);

                            System.Collections.Specialized.StringCollection Lista1 = new System.Collections.Specialized.StringCollection();
                            foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                            {
                                if (Combo1 is System.Windows.Forms.ComboBox)
                                {
                                    String Text1 = Combo1.Text;
                                    if (Text1 != "")
                                    {
                                        if (Lista1.Contains(Text1) == false)
                                        {
                                            Lista1.Add(Text1);
                                        }
                                    }
                                }
                            }

                            if (Lista1.Count > 0)
                            {
                                ObjectIdCollection Draw_order_Ids = new ObjectIdCollection();
                                foreach (ObjectId O_id in BTrecord)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(O_id, OpenMode.ForRead);
                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                    {
                                        Polyline Poly1 = (Polyline)Ent1;
                                        if (Lista1.Contains(Poly1.Layer) == true)
                                        {
                                            Poly1.UpgradeOpen();
                                            if (Poly1.Closed == true)
                                            {
                                                Poly1.ColorIndex = 7;
                                            }
                                            else
                                            {
                                                Poly1.ColorIndex = 1;
                                                Draw_order_Ids.Add(Poly1.ObjectId);
                                            }
                                        }
                                    }
                                }
                                Autodesk.AutoCAD.DatabaseServices.DrawOrderTable DrawOrderTable1 = (Autodesk.AutoCAD.DatabaseServices.DrawOrderTable)Trans1.GetObject(BTrecord.DrawOrderTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                if (Draw_order_Ids.Count > 0)
                                {
                                    DrawOrderTable1.MoveToTop(Draw_order_Ids);
                                }
                            }
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Editor1.WriteMessage("\nCommand:");
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
            }
        }

        private void Button_dimension_To_CL_Click(object sender, EventArgs e)
        {

            // Private Sub Button_dimension_To_CL_Click(sender As Object, e As EventArgs) Handles Button_dimension_To_CL.Click


            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                String No_plot = "NO PLOT";
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;


                Editor1.SetImpliedSelection(Empty_array);
                try
                {
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();



                    Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat1;

                    Autodesk.AutoCAD.EditorInput.PromptEntityOptions Object_Prompt = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect centerline:");

                    Object_Prompt.SetRejectMessage("\nPlease select a lightweight polyline");
                    Object_Prompt.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), true);


                    Rezultat1 = Editor1.GetEntity(Object_Prompt);


                    if (Rezultat1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        Editor1.WriteMessage("\nCommand:");
                        Freeze_operations = false;
                        Editor1.SetImpliedSelection(Empty_array);
                        return;
                    }
                    if (Rezultat1.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                    {
                        if (Rezultat1 != null)
                        {
                            using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                            {

                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Polyline PolyCL = (Polyline)Trans1.GetObject(Rezultat1.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);



                                    if (PolyCL.Elevation != 0)
                                    {
                                        Freeze_operations = false;
                                        MessageBox.Show("CL Polyline is not at elevation 0");
                                        return;

                                    }
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                                    Functions.Creaza_layer(No_plot, 40, false);

                                    System.Collections.Specialized.StringCollection Lista1 = new System.Collections.Specialized.StringCollection();
                                    foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                                    {
                                        if (Combo1 is System.Windows.Forms.ComboBox)
                                        {
                                            String Text1 = Combo1.Text;
                                            if (Text1 != "")
                                            {
                                                if (Lista1.Contains(Text1) == false)
                                                {
                                                    Lista1.Add(Text1);
                                                }
                                            }
                                        }
                                    }
                                    if (Lista1.Count > 0)
                                    {
                                        label_processing2.Visible = true;
                                        this.Refresh();
                                        foreach (ObjectId objID in BTrecord)
                                        {
                                            Entity Ent1 = (Entity)Trans1.GetObject(objID, OpenMode.ForRead);
                                            if (Ent1 is Autodesk.AutoCAD.DatabaseServices.RotatedDimension & Ent1.Layer == No_plot)
                                            {
                                                Ent1.UpgradeOpen();
                                                Ent1.Erase();
                                            }
                                        }


                                        foreach (ObjectId objID in BTrecord)
                                        {
                                            try
                                            {
                                                Entity Ent1 = (Entity)Trans1.GetObject(objID, OpenMode.ForRead);

                                                if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline & objID != PolyCL.ObjectId & Lista1.Contains(Ent1.Layer) == true)
                                                {
                                                    Polyline Poly1 = (Polyline)Ent1;
                                                    if (Poly1.Closed == true)
                                                    {


                                                        Poly1.UpgradeOpen();

                                                        Poly1.ColorIndex = 7;
                                                        for (int i = 0; i <= Poly1.NumberOfVertices - 2; ++i)
                                                        {
                                                            Point3d Mid_point_poly1 = new Point3d();
                                                            Mid_point_poly1 = Poly1.GetPointAtParameter(i + 0.5);
                                                            Double Bearing1 = Math.Round(Functions.GET_Bearing_rad(Poly1.GetPointAtParameter(i).X, Poly1.GetPointAtParameter(i).Y, Poly1.GetPointAtParameter(i + 1).X, Poly1.GetPointAtParameter(i + 1).Y), 4);
                                                            Point3d Pt_on_poly = new Point3d();
                                                            Pt_on_poly = PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true);
                                                            Double Paramcl = PolyCL.GetParameterAtPoint(Pt_on_poly);

                                                            Point3d PT_i = new Point3d();
                                                            PT_i = PolyCL.GetPointAtParameter(Math.Floor(Paramcl));

                                                            Point3d PT_i_1 = new Point3d();
                                                            PT_i_1 = PolyCL.GetPointAtParameter(Math.Ceiling(Paramcl));

                                                            if (Math.Floor(Paramcl) == Math.Ceiling(Paramcl))
                                                            {
                                                                if (Math.Floor(Paramcl) >= 1 & Math.Floor(Paramcl) < PolyCL.NumberOfVertices - 1)
                                                                {
                                                                    PT_i_1 = PolyCL.GetPointAtParameter(Math.Floor(Paramcl) + 1);
                                                                }
                                                                else
                                                                {
                                                                    PT_i = PolyCL.GetPointAtParameter(0);
                                                                    PT_i_1 = PolyCL.GetPointAtParameter(1);
                                                                }
                                                            }




                                                            Double Bearing_cl = Math.Round(Functions.GET_Bearing_rad(PT_i.X, PT_i.Y, PT_i_1.X, PT_i_1.Y), 4);

                                                            do
                                                            {
                                                                Bearing_cl = Math.Round(Bearing_cl - Math.Round(Math.PI, 4), 4);

                                                            } while (Bearing_cl >= Math.Round(Math.PI, 4));

                                                            do
                                                            {
                                                                Bearing1 = Math.Round(Bearing1 - Math.Round(Math.PI, 4), 4);

                                                            } while (Bearing1 >= Math.Round(Math.PI, 4));

                                                            if (Bearing_cl == Bearing1 | Bearing_cl + Math.Round(Math.PI, 4) == Bearing1 | Bearing_cl - Math.Round(Math.PI, 4) == Bearing1)
                                                            {
                                                                Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true));
                                                                Line1.Layer = No_plot;
                                                                Double Offset = Line1.Length;
                                                                Double a = Math.Round(Offset, 3);
                                                                Double b = Math.Round(Offset, 0);
                                                                if (a != b)
                                                                {
                                                                    RotatedDimension Dimension1 = new RotatedDimension();
                                                                    Dimension1.Layer = No_plot;
                                                                    Dimension1.XLine1Point = Line1.StartPoint;
                                                                    Dimension1.XLine2Point = Line1.EndPoint;
                                                                    Dimension1.Rotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                    Dimension1.DimLinePoint = Line1.StartPoint;
                                                                    Dimension1.UsingDefaultTextPosition = true;
                                                                    Dimension1.TextAttachment = AttachmentPoint.MiddleCenter;
                                                                    Dimension1.TextRotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                    Dimension1.Dimasz = 2; //'Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                    Dimension1.Dimdec = 4; //'Sets the number of decimal places displayed for the primary units of a dimension
                                                                    Dimension1.Dimtxt = 4; //'Specif (ies the height of dimension text, unless the current text style has a fixed height
                                                                    Functions.add_extra_param_to_dim(Dimension1, ThisDrawing);
                                                                    BTrecord.AppendEntity(Dimension1);
                                                                    Trans1.AddNewlyCreatedDBObject(Dimension1, true);
                                                                }
                                                            }
                                                            else
                                                            {

                                                                Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true));
                                                                Line1.Layer = No_plot;

                                                                Double L1 = Line1.Length;

                                                                Double a = Math.Round((Bearing1 + Math.Round(Math.PI, 4) / 2), 4);
                                                                Double b = Math.Round(Bearing_cl, 4);
                                                                Double c = Math.Round((Bearing1 - Math.Round(Math.PI, 4) / 2), 4);
                                                                Double d = Math.Round(Bearing1, 4);
                                                                bool Add_L = false;

                                                                if (a != b & c != b & d != b)
                                                                {
                                                                    Add_L = true;
                                                                }

                                                                if (Math.Round(L1, 0) == Math.Round(L1, 3))
                                                                {
                                                                    Add_L = false;
                                                                }


                                                                if (Add_L == true)
                                                                {
                                                                    RotatedDimension Dimension1 = new RotatedDimension();
                                                                    Dimension1.Layer = No_plot;
                                                                    Dimension1.XLine1Point = Line1.StartPoint;
                                                                    Dimension1.XLine2Point = Line1.EndPoint;
                                                                    Dimension1.Rotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                    Dimension1.DimLinePoint = Line1.StartPoint;
                                                                    Dimension1.UsingDefaultTextPosition = true;
                                                                    Dimension1.TextAttachment = AttachmentPoint.MiddleCenter;
                                                                    Dimension1.TextRotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                    Dimension1.Dimasz = 2; //Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                    Dimension1.Dimdec = 4; //Sets the number of decimal places displayed for the primary units of a dimension
                                                                    Dimension1.Dimtxt = 4; //Specif (ies the height of dimension text, unless the current text style has a fixed height
                                                                    Functions.add_extra_param_to_dim(Dimension1, ThisDrawing);
                                                                    BTrecord.AppendEntity(Dimension1);
                                                                    Trans1.AddNewlyCreatedDBObject(Dimension1, true);
                                                                }


                                                            }


                                                        }

                                                        if (Poly1.Closed == true)
                                                        {
                                                            Point3d Mid_point_poly1 = new Point3d();
                                                            Autodesk.AutoCAD.DatabaseServices.Line Line1_ws = new Autodesk.AutoCAD.DatabaseServices.Line(Poly1.GetPointAtDist(0), Poly1.GetPointAtParameter(Poly1.NumberOfVertices - 1));
                                                            if (Line1_ws.Length > 0.01)
                                                            {
                                                                Mid_point_poly1 = Line1_ws.GetPointAtDist(Line1_ws.Length / 2);
                                                                Double Bearing1 = Math.Round(Functions.GET_Bearing_rad(Line1_ws.StartPoint.X, Line1_ws.StartPoint.Y, Line1_ws.EndPoint.X, Line1_ws.EndPoint.Y), 4);
                                                                Point3d Pt_on_poly = new Point3d();
                                                                Pt_on_poly = PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true);
                                                                Double Paramcl = PolyCL.GetParameterAtPoint(Pt_on_poly);
                                                                Point3d PT_i = new Point3d();
                                                                PT_i = PolyCL.GetPointAtParameter(Math.Floor(Paramcl));

                                                                Point3d PT_i_1 = new Point3d();
                                                                PT_i_1 = PolyCL.GetPointAtParameter(Math.Ceiling(Paramcl));

                                                                if (Math.Floor(Paramcl) == Math.Ceiling(Paramcl))
                                                                {
                                                                    if (Math.Floor(Paramcl) >= 1 & Math.Floor(Paramcl) < PolyCL.NumberOfVertices - 1)
                                                                    {
                                                                        PT_i_1 = PolyCL.GetPointAtParameter(Math.Floor(Paramcl) + 1);
                                                                    }
                                                                    else
                                                                    {
                                                                        PT_i = PolyCL.GetPointAtParameter(0);
                                                                        PT_i_1 = PolyCL.GetPointAtParameter(1);
                                                                    }
                                                                }


                                                                Double Bearing_cl = Math.Round(Functions.GET_Bearing_rad(PT_i.X, PT_i.Y, PT_i_1.X, PT_i_1.Y), 4);
                                                                do
                                                                {
                                                                    Bearing_cl = Math.Round(Bearing_cl - Math.Round(Math.PI, 4), 4);
                                                                } while (Bearing_cl >= Math.Round(Math.PI, 4));


                                                                do
                                                                {
                                                                    Bearing1 = Math.Round(Bearing1 - Math.Round(Math.PI, 4), 4);
                                                                } while (Bearing1 >= Math.Round(Math.PI, 4));

                                                                if (Bearing_cl == Bearing1 | Bearing_cl + Math.Round(Math.PI, 4) == Bearing1 | Bearing_cl - Math.Round(Math.PI, 4) == Bearing1)
                                                                {
                                                                    Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true));
                                                                    Line1.Layer = No_plot;
                                                                    Double Offset = Line1.Length;
                                                                    Double a = Math.Round(Offset, 3);
                                                                    Double b = Math.Round(Offset, 0);
                                                                    if (a != b)
                                                                    {
                                                                        RotatedDimension Dimension1 = new RotatedDimension();
                                                                        Dimension1.Layer = No_plot;
                                                                        Dimension1.XLine1Point = Line1.StartPoint;
                                                                        Dimension1.XLine2Point = Line1.EndPoint;
                                                                        Dimension1.Rotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                        Dimension1.DimLinePoint = Line1.StartPoint;
                                                                        Dimension1.UsingDefaultTextPosition = true;
                                                                        Dimension1.TextAttachment = AttachmentPoint.MiddleCenter;
                                                                        Dimension1.TextRotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                        Dimension1.Dimasz = 2; //Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                        Dimension1.Dimdec = 4; //Sets the number of decimal places displayed for the primary units of a dimension
                                                                        Dimension1.Dimtxt = 4; //Specif (ies the height of dimension text, unless the current text style has a fixed height
                                                                        Functions.add_extra_param_to_dim(Dimension1, ThisDrawing);
                                                                        BTrecord.AppendEntity(Dimension1);
                                                                        Trans1.AddNewlyCreatedDBObject(Dimension1, true);
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    Autodesk.AutoCAD.DatabaseServices.Line Line1 = new Autodesk.AutoCAD.DatabaseServices.Line(Mid_point_poly1, PolyCL.GetClosestPointTo(Mid_point_poly1, Vector3d.ZAxis, true));
                                                                    Line1.Layer = No_plot;
                                                                    Double L1 = Line1.Length;
                                                                    Double a = Math.Round((Bearing1 + Math.Round(Math.PI, 4) / 2), 4);
                                                                    Double b = Math.Round(Bearing_cl, 4);
                                                                    Double c = Math.Round((Bearing1 - Math.Round(Math.PI, 4) / 2), 4);
                                                                    Double d = Math.Round(Bearing1, 4);
                                                                    bool Add_L = false;
                                                                    if (a != b & c != b & d != b)
                                                                    {
                                                                        Add_L = true;
                                                                    }
                                                                    if (Math.Round(L1, 0) == Math.Round(L1, 3))
                                                                    {
                                                                        Add_L = false;
                                                                    }
                                                                    if (Add_L == true)
                                                                    {
                                                                        RotatedDimension Dimension1 = new RotatedDimension();
                                                                        Dimension1.Layer = No_plot;
                                                                        Dimension1.XLine1Point = Line1.StartPoint;
                                                                        Dimension1.XLine2Point = Line1.EndPoint;
                                                                        Dimension1.Rotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                        Dimension1.DimLinePoint = Line1.StartPoint;
                                                                        Dimension1.UsingDefaultTextPosition = true;
                                                                        Dimension1.TextAttachment = AttachmentPoint.MiddleCenter;
                                                                        Dimension1.TextRotation = Functions.GET_Bearing_rad(Line1.StartPoint.X, Line1.StartPoint.Y, Line1.EndPoint.X, Line1.EndPoint.Y);
                                                                        Dimension1.Dimasz = 2;//Controls the size of dimension line and leader line arrowheads. Also controls the size of hook lines
                                                                        Dimension1.Dimdec = 4; //Sets the number of decimal places displayed for the primary units of a dimension
                                                                        Dimension1.Dimtxt = 4; //Specif (ies the height of dimension text, unless the current text style has a fixed height
                                                                        Functions.add_extra_param_to_dim(Dimension1, ThisDrawing);
                                                                        BTrecord.AppendEntity(Dimension1);
                                                                        Trans1.AddNewlyCreatedDBObject(Dimension1, true);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Poly1.UpgradeOpen();
                                                        Poly1.ColorIndex = 1;
                                                    }
                                                }

                                            }
                                            catch (System.Exception ex)
                                            {

                                            }
                                        }

                                        Trans1.Commit();
                                    }
                                }
                            }
                        }
                    }
                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                }
                catch (System.Exception ex)
                {
                    Editor1.WriteMessage("\nCommand:");
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
                label_processing2.Visible = false;
            }

        }

        private void Button_analise_gaps_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                String No_plot = "NO PLOT";
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Editor1.SetImpliedSelection(Empty_array);

                int no_regions = 0;

                try
                {
                    using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                            DBObjectCollection Colectie_poly_closed = new DBObjectCollection();
                            Functions.Creaza_layer(No_plot, 40, false);
                            System.Collections.Specialized.StringCollection Lista_layere = new System.Collections.Specialized.StringCollection();

                            foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                            {
                                if (Combo1 is System.Windows.Forms.ComboBox)
                                {
                                    String Text1 = Combo1.Text;
                                    if (Text1 != "")
                                    {
                                        if (Lista_layere.Contains(Text1) == false)
                                        {
                                            Lista_layere.Add(Text1);
                                        }
                                    }
                                }
                            }

                            if (Lista_layere.Count > 0)
                            {
                                label_processing2.Visible = true;
                                this.Refresh();
                                foreach (ObjectId O_id in BTrecord)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(O_id, OpenMode.ForRead);
                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Region & Ent1.Layer == No_plot)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Region Region0 = (Autodesk.AutoCAD.DatabaseServices.Region)Ent1;
                                        Region0.UpgradeOpen();
                                        Region0.Erase();
                                    }
                                }

                                DBObjectCollection Poly_colection = new DBObjectCollection();
                                Autodesk.AutoCAD.DatabaseServices.Region Region1 = new Autodesk.AutoCAD.DatabaseServices.Region();


                                bool IS_first = true;
                                foreach (ObjectId O_id in BTrecord)
                                {
                                    Entity Ent1 = (Entity)Trans1.GetObject(O_id, OpenMode.ForRead);
                                    if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                    {
                                        Polyline Poly1 = (Polyline)Ent1;
                                        if (Lista_layere.Contains(Poly1.Layer) == true)
                                        {
                                            Poly1.UpgradeOpen();

                                            if (Poly1.Closed == true)
                                            {


                                                try
                                                {
                                                    Autodesk.AutoCAD.DatabaseServices.Region Region0 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                                    DBObjectCollection Poly_Colection0 = new DBObjectCollection();
                                                    Poly_Colection0.Add(Poly1);
                                                    DBObjectCollection Region_Colection0 = new DBObjectCollection();
                                                    Region_Colection0 = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection0);
                                                    Region0 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection0[0];

                                                    Poly1.ColorIndex = 7;
                                                    Poly_colection.Add(Poly1);

                                                    if (IS_first == true)
                                                    {

                                                        DBObjectCollection Poly_Colection1 = new DBObjectCollection();
                                                        Poly_Colection1.Add(Poly1);
                                                        DBObjectCollection Region_Colection1 = new DBObjectCollection();
                                                        Region_Colection1 = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection1);
                                                        Region1 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[0];

                                                        Region1.ColorIndex = 5;
                                                        Region1.Layer = No_plot;
                                                        BTrecord.AppendEntity(Region1);
                                                        Trans1.AddNewlyCreatedDBObject(Region1, true);
                                                        IS_first = false;
                                                        no_regions = no_regions + 1;
                                                    }



                                                }
                                                catch (System.Exception ex)
                                                {
                                                    Poly1.ColorIndex = 1;
                                                }
                                            }
                                            else
                                            {
                                                Poly1.ColorIndex = 1;
                                            }
                                        }
                                    }

                                }



                                if (Poly_colection.Count > 0)
                                {

                                    int startB = 0;







                                    for (int i = 1; i < Poly_colection.Count; ++i)
                                    {
                                        try
                                        {

                                            DBObjectCollection Poly_collection1 = new DBObjectCollection();
                                            Poly_collection1.Add(Poly_colection[i]);
                                            DBObjectCollection Region_Collection1 = new DBObjectCollection();
                                            Region_Collection1 = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_collection1);
                                            Autodesk.AutoCAD.DatabaseServices.Region Region2 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Collection1[0];
                                            Region2.ColorIndex = 5;

                                            Region1.BooleanOperation(BooleanOperationType.BoolUnite, Region2);
                                            Region1.ColorIndex = 5;

                                        }
                                        catch (System.Exception ex)
                                        {


                                            Polyline Poly1b = (Polyline)Poly_colection[i];
                                            Poly1b.UpgradeOpen();
                                            Poly1b.ColorIndex = 1;

                                            DBObjectCollection Poly_collection1B = new DBObjectCollection();
                                            Poly_collection1B.Add(Poly_colection[startB]);


                                            DBObjectCollection Region_collection1B = new DBObjectCollection();
                                            Region_collection1B = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_collection1B);

                                            Autodesk.AutoCAD.DatabaseServices.Region Region1b = new Autodesk.AutoCAD.DatabaseServices.Region();
                                            Region1b = (Autodesk.AutoCAD.DatabaseServices.Region)Region_collection1B[0];

                                            Region1b.ColorIndex = 5;
                                            Region1b.Layer = No_plot;
                                            BTrecord.AppendEntity(Region1b);
                                            Trans1.AddNewlyCreatedDBObject(Region1b, true);
                                            no_regions = no_regions + 1;

                                            for (int j = startB + 1; j < i; ++j)
                                            {
                                                DBObjectCollection Poly_collection2B = new DBObjectCollection();
                                                Poly_collection2B.Add(Poly_colection[j]);

                                                DBObjectCollection Region_collection2B = new DBObjectCollection();
                                                Region_collection2B = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_collection2B);


                                                Autodesk.AutoCAD.DatabaseServices.Region Region2b = (Autodesk.AutoCAD.DatabaseServices.Region)Region_collection2B[0];
                                                Region2b.ColorIndex = 5;

                                                Region1b.BooleanOperation(BooleanOperationType.BoolUnite, Region2b);
                                                Region1b.ColorIndex = 5;
                                            }
                                            startB = i + 1;

                                        }
                                    }


                                }





                                Trans1.Commit();
                            }
                        }
                    }

                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                }
                catch (System.Exception ex)
                {
                    Editor1.WriteMessage("\nCommand:");
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;

                string Plural = "s";
                if (no_regions == 1) Plural = "";
                MessageBox.Show(no_regions.ToString() + " region" + Plural + " created");
                label_processing2.Visible = false;
            }
            //End Sub
        }

        private void Button_analise_overlapp_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show("The processing time depends on the number of polylines existing in the layers you selected" + vbcrlf + "(Example: takes around 90 seconds to process 300 polylines)" +
                                vbcrlf + "Do you want to continue?", "WARNING", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }


            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                String No_plot = "NO PLOT";
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Editor1.SetImpliedSelection(Empty_array);
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                try
                {
                    using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                            DBObjectCollection Poly_colection = new DBObjectCollection();

                            Functions.Creaza_layer(No_plot, 40, false);
                            System.Collections.Specialized.StringCollection Lista1 = new System.Collections.Specialized.StringCollection();

                            foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                            {
                                if (Combo1 is System.Windows.Forms.ComboBox)
                                {
                                    String Text1 = Combo1.Text;
                                    if (Text1 != "")
                                    {
                                        if (Lista1.Contains(Text1) == false)
                                        {
                                            Lista1.Add(Text1);
                                        }
                                    }
                                }
                            }

                            DBObjectCollection Region_Colection1 = new DBObjectCollection();

                            if (Lista1.Count > 0)
                            {
                                label_processing2.Visible = true;
                                this.Refresh();

                                foreach (ObjectId O_id in BTrecord)
                                {
                                    using (Entity Ent1 = (Entity)Trans1.GetObject(O_id, OpenMode.ForRead))
                                    {
                                        if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                        {
                                            Polyline Poly1 = (Polyline)Ent1;
                                            if (Lista1.Contains(Poly1.Layer) == true)
                                            {
                                                Poly1.UpgradeOpen();

                                                if (Poly1.Closed == true)
                                                {
                                                    try
                                                    {
                                                        Autodesk.AutoCAD.DatabaseServices.Region Region0 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                                        DBObjectCollection Poly_Colection0 = new DBObjectCollection();
                                                        Poly_Colection0.Add(Poly1);

                                                        DBObjectCollection Region_Colection0 = new DBObjectCollection();
                                                        Region_Colection0 = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(Poly_Colection0);
                                                        Region0 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection0[0];
                                                        Region_Colection1.Add(Region0);




                                                        Poly1.ColorIndex = 7;
                                                        Poly_colection.Add(Poly1);
                                                    }
                                                    catch (System.Exception ex)
                                                    {
                                                        Poly1.ColorIndex = 1;
                                                    }
                                                }
                                                else
                                                {
                                                    Poly1.ColorIndex = 1;
                                                }
                                            }
                                        }

                                        if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Region & Ent1.Layer == No_plot)
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.Region Region1 = (Autodesk.AutoCAD.DatabaseServices.Region)Ent1;
                                            Region1.UpgradeOpen();
                                            Region1.Erase();
                                        }

                                    }

                                }

                                if (Poly_colection.Count > 1)
                                {
                                    for (int i = 0; i < Poly_colection.Count; ++i)
                                    {
                                        Autodesk.AutoCAD.DatabaseServices.Region Region1 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                        Region1 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[i].Clone();

                                        for (int j = i + 1; j < Poly_colection.Count; ++j)
                                        {
                                            Autodesk.AutoCAD.DatabaseServices.Region Region2 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                            Region2 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[j].Clone();
                                            try
                                            {
                                                Region1.BooleanOperation(BooleanOperationType.BoolIntersect, Region2);
                                                if (Region1.Area > 0.001)
                                                {
                                                    Region1.ColorIndex = 1;
                                                    Region1.Layer = No_plot;
                                                    BTrecord.AppendEntity(Region1);
                                                    Trans1.AddNewlyCreatedDBObject(Region1, true);
                                                }
                                                Region1 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                                Region1 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[i].Clone();
                                            }
                                            catch (System.Exception ex)
                                            {
                                                Region1 = new Autodesk.AutoCAD.DatabaseServices.Region();
                                                Region1 = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[i].Clone();
                                                Region1.ColorIndex = 1;
                                                Region1.Layer = No_plot;
                                                BTrecord.AppendEntity(Region1);
                                                Trans1.AddNewlyCreatedDBObject(Region1, true);

                                                Autodesk.AutoCAD.DatabaseServices.Region Region2b = new Autodesk.AutoCAD.DatabaseServices.Region();
                                                Region2b = (Autodesk.AutoCAD.DatabaseServices.Region)Region_Colection1[j].Clone();
                                                Region2b.ColorIndex = 1;
                                                Region2b.Layer = No_plot;
                                                BTrecord.AppendEntity(Region2b);
                                                Trans1.AddNewlyCreatedDBObject(Region2b, true);

                                            }

                                        }
                                    }

                                }
                            }

                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Editor1.WriteMessage("\nCommand:");
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
                label_processing2.Visible = false;
            }
        }

        private void button_topology_Click(object sender, EventArgs e)
        {
            bool Runn = true;



            if (Freeze_operations == false & Runn == true)
            {
                Freeze_operations = true;


                String No_plot = "NO PLOT";
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Editor1.SetImpliedSelection(Empty_array);
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                try
                {
                    using (DocumentLock Lock_dwg = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (Autodesk.AutoCAD.DatabaseServices.BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                            DBObjectCollection Poly_colection = new DBObjectCollection();

                            Functions.Creaza_layer(No_plot, 40, false);
                            System.Collections.Specialized.StringCollection Lista1 = new System.Collections.Specialized.StringCollection();
                            foreach (System.Windows.Forms.Control Combo1 in panel_layerpanel.Controls)
                            {
                                if (Combo1 is System.Windows.Forms.ComboBox)
                                {
                                    String Text1 = Combo1.Text;
                                    if (Text1 != "")
                                    {
                                        if (Lista1.Contains(Text1) == false)
                                        {
                                            Lista1.Add(Text1);
                                        }
                                    }
                                }
                            }



                            if (Lista1.Count > 0)
                            {
                                label_processing2.Visible = true;
                                this.Refresh();
                                ObjectIdCollection ids1 = new ObjectIdCollection();
                                int index1 = 0;

                                foreach (ObjectId O_id in BTrecord)
                                {
                                    using (Entity Ent1 = (Entity)Trans1.GetObject(O_id, OpenMode.ForRead))
                                    {
                                        if (Ent1 is Autodesk.AutoCAD.DatabaseServices.Polyline)
                                        {
                                            Polyline Poly1 = (Polyline)Ent1;
                                            if (Lista1.Contains(Poly1.Layer) == true)
                                            {
                                                if (index1 == 0)
                                                {
                                                    ids1.Add(O_id);
                                                }
                                                index1 = index1 + 1;
                                            }
                                        }



                                    }

                                }



                                //Autodesk.Gis.Map.Topology.EntityCreationSettings.Create(Autodesk.Gis.Map.Topology.TopologyTypes.Polygon, IntPtr.Zero, true);

                                Autodesk.Gis.Map.Topology.CreateOptions options = Autodesk.Gis.Map.Topology.CreateOptions.HighlightSliverPolygons;
                                Autodesk.Gis.Map.Topology.Topologies Topos1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.Topologies;
                                Topos1.Create("aa", ids1, Autodesk.Gis.Map.Topology.TopologyTypes.Polygon);


                                Autodesk.Gis.Map.Topology.TopologyModel Tm1 = Topos1["aa"];
                                Tm1.Open(Autodesk.Gis.Map.Topology.OpenMode.ForRead);




                                if (Poly_colection.Count > 1)
                                {
                                    for (int i = 0; i < Poly_colection.Count; ++i)
                                    {


                                    }

                                }
                            }

                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Editor1.WriteMessage("\nCommand:");
                    MessageBox.Show(ex.Message);
                }
                Freeze_operations = false;
                label_processing2.Visible = false;
            }
        }

        private void button_browse_select_output_folder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_output_folder.Text = fbd.SelectedPath.ToString();
                }

            }
        }

        private void button_load_layers_Click(object sender, EventArgs e)
        {
            List<string> lista1 = get_layers_from_dwg();
            dt_layer = new System.Data.DataTable();
            dt_layer.Columns.Add("Select", typeof(bool));
            dt_layer.Columns.Add("Name", typeof(string));
            dt_layer.Columns.Add("Export as Polygon", typeof(bool));

            for (int i = 0; i < lista1.Count; ++i)
            {
                dt_layer.Rows.Add();
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Select"] = false;
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Name"] = lista1[i];
                dt_layer.Rows[dt_layer.Rows.Count - 1]["Export as Polygon"] = false;

            }

            if (lista1.Count > 0)
            {
                dataGridView_prop.DataSource = dt_layer;
                dataGridView_prop.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_prop.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_prop.EnableHeadersVisualStyles = false;
            }
        }

        public void copy_od(ObjectId id1, ObjectId id2)
        {
            System.Data.DataTable Data_table_for_object_data = new System.Data.DataTable();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {
                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                    Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                    Autodesk.Gis.Map.ObjectData.Records Records1;
                    Autodesk.Gis.Map.ObjectData.Records Records2;

                    Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                    Entity Ent2 = Trans1.GetObject(id2, OpenMode.ForRead) as Entity;

                    if (Ent1 != null & Ent2 != null & Ent1.ObjectId != Ent2.ObjectId)
                    {
                        try
                        {
                            using (Records2 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent2.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForWrite, false))
                            {
                                if (Records2 != null)
                                {
                                    if (Records2.Count > 0)
                                    {
                                        System.Collections.IEnumerator ie = Records2.GetEnumerator();
                                        while (ie.MoveNext())
                                        {
                                            Records2.RemoveRecord();
                                        }
                                    }
                                }
                            }

                            using (Records1 = Tables1.GetObjectRecords(Convert.ToUInt32(0), Ent1.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                            {
                                if (Records1 != null)
                                {
                                    if (Records1.Count > 0)
                                    {
                                        Data_table_for_object_data.Rows.Add();
                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                        {
                                            Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Record1.TableName];
                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                            Tabla1.AddRecord(Record1, Ent2.ObjectId);
                                            for (int j = 0; j < Record1.Count; ++j)
                                            {
                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[j];
                                                string Nume_field = Field_def1.Name;
                                                string Valoare_field = (string)Record1[j].StrValue;
                                                if (Data_table_for_object_data.Columns.Contains(Nume_field) == false)
                                                {
                                                    Data_table_for_object_data.Columns.Add(Nume_field, typeof(String));
                                                }
                                                Data_table_for_object_data.Rows[0][Nume_field] = Valoare_field;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (AccessViolationException ex1)
                        {
                            MessageBox.Show(ex1.Message);
                        }
                    }
                    Trans1.Commit();
                }
            }
        }

        private void button_map_export_to_shp_Click(object sender, EventArgs e)
        {
            if (dt_layer != null && dt_layer.Rows.Count > 0)
            {
                List<string> lista_selected = new List<string>();
                List<string> lista_selected_user = new List<string>();
                List<string> lista_polygon_layers = new List<string>();
                List<string> lista_polygon_layers_user = new List<string>();
                for (int i = 0; i < dt_layer.Rows.Count; ++i)
                {
                    if ((bool)dt_layer.Rows[i][0] == true && (bool)dt_layer.Rows[i][2] == false)
                    {
                        lista_selected.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }
                    if ((bool)dt_layer.Rows[i][0] == true && (bool)dt_layer.Rows[i][2] == true)
                    {
                        lista_polygon_layers.Add(Convert.ToString(dt_layer.Rows[i][1]));
                    }
                }

                if (lista_selected.Count > 0 || lista_polygon_layers.Count > 0)
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
                                BlockTableRecord BTrecord_MS = Trans1.GetObject(BlockTable1[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                                BlockTableRecord BTrecord_PS = Trans1.GetObject(BlockTable1[BlockTableRecord.PaperSpace], OpenMode.ForRead) as BlockTableRecord;
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                                TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                                Autodesk.Gis.Map.MapApplication mapApp = Autodesk.Gis.Map.HostMapApplicationServices.Application;

                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                dt1.Columns.Add("id", typeof(ObjectId));
                                dt1.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add("id", typeof(ObjectId));
                                dt2.Columns.Add("layer", typeof(string));

                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                dt3.Columns.Add("id", typeof(ObjectId));
                                dt3.Columns.Add("layer", typeof(string));

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_rez.MessageForAdding = "\nSelect the objects:";
                                Prompt_rez.SingleOnly = false;

                                bool open_poly_in_Polygon_layer = false;
                                bool linie_in_Polygon_layer = false;
                                bool point_in_line_layer = false;


                                List<string> lista_polylines_layers = new List<string>();
                                List<string> lista_point_layers = new List<string>();


                                if (checkBox_user_select_export.Checked == true)
                                {
                                    #region checkBox_user_select_export
                                    Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                                    if (Rezultat1.Status == PromptStatus.OK)

                                    {
                                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                                        {

                                            Entity Ent1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as Entity;


                                            if (lista_polygon_layers.Contains(Ent1.Layer) == true)
                                            {
                                                if (Ent1 is Polyline)
                                                {
                                                    Polyline poly1 = Ent1 as Polyline;
                                                    if (poly1.Closed == true)
                                                    {
                                                        dt3.Rows.Add();
                                                        dt3.Rows[dt3.Rows.Count - 1][0] = Ent1.ObjectId;
                                                        dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                                        if (lista_polygon_layers_user.Contains(Ent1.Layer) == false) lista_polygon_layers_user.Add(Ent1.Layer);

                                                    }
                                                    else
                                                    {
                                                        Ent1.UpgradeOpen();
                                                        Ent1.ColorIndex = 42;
                                                        open_poly_in_Polygon_layer = true;
                                                    }
                                                }
                                                else if (Ent1 is MPolygon)
                                                {
                                                    MPolygon mpolyg1 = Ent1 as MPolygon;
                                                    dt3.Rows.Add();
                                                    dt3.Rows[dt3.Rows.Count - 1][0] = Ent1.ObjectId;
                                                    dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                                    if (lista_polygon_layers_user.Contains(Ent1.Layer) == false) lista_polygon_layers_user.Add(Ent1.Layer);

                                                }
                                                else
                                                {
                                                    Ent1.UpgradeOpen();
                                                    Ent1.ColorIndex = 42;
                                                    linie_in_Polygon_layer = true;
                                                }
                                            }



                                            if (Ent1 is Curve && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = Rezultat1.Value[i].ObjectId;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_selected_user.Contains(Ent1.Layer) == false) lista_selected_user.Add(Ent1.Layer);

                                            }
                                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_selected.Contains(Ent1.Layer) == true)
                                            {
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = Rezultat1.Value[i].ObjectId;
                                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_selected_user.Contains(Ent1.Layer) == false) lista_selected_user.Add(Ent1.Layer);

                                            }



                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        Entity Ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;

                                        if (lista_polygon_layers.Contains(Ent1.Layer) == true)
                                        {
                                            if (Ent1 is Polyline)
                                            {
                                                Polyline poly1 = Ent1 as Polyline;
                                                if (poly1.Closed == true)
                                                {
                                                    dt3.Rows.Add();
                                                    dt3.Rows[dt3.Rows.Count - 1][0] = id1;
                                                    dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                                }
                                                else
                                                {
                                                    Ent1.UpgradeOpen();
                                                    Ent1.ColorIndex = 42;
                                                    open_poly_in_Polygon_layer = true;
                                                }
                                            }
                                            else if (Ent1 is MPolygon)
                                            {
                                                MPolygon mpolyg1 = Ent1 as MPolygon;
                                                dt3.Rows.Add();
                                                dt3.Rows[dt3.Rows.Count - 1][0] = id1;
                                                dt3.Rows[dt3.Rows.Count - 1][1] = Ent1.Layer;
                                            }
                                            else
                                            {
                                                Ent1.UpgradeOpen();
                                                Ent1.ColorIndex = 42;
                                                linie_in_Polygon_layer = true;
                                            }
                                        }

                                        if (lista_selected.Contains(Ent1.Layer) == true)
                                        {
                                            if (Ent1 is Curve && lista_point_layers.Contains(Ent1.Layer) == false)
                                            {
                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][0] = id1;
                                                dt1.Rows[dt1.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_polylines_layers.Contains(Ent1.Layer) == false) lista_polylines_layers.Add(Ent1.Layer);
                                            }
                                            else if (Ent1 is Curve && lista_point_layers.Contains(Ent1.Layer) == true)
                                            {
                                                Ent1.UpgradeOpen();
                                                Ent1.ColorIndex = 42;
                                                point_in_line_layer = true;
                                            }

                                            if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_polylines_layers.Contains(Ent1.Layer) == false)
                                            {
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = id1;
                                                dt2.Rows[dt2.Rows.Count - 1][1] = Ent1.Layer;
                                                if (lista_point_layers.Contains(Ent1.Layer) == false) lista_point_layers.Add(Ent1.Layer);
                                            }
                                            else if ((Ent1 is DBPoint || Ent1 is BlockReference) && lista_polylines_layers.Contains(Ent1.Layer) == true)
                                            {
                                                Ent1.UpgradeOpen();
                                                Ent1.ColorIndex = 42;
                                                point_in_line_layer = true;
                                            }

                                        }

                                    }
                                }



                                if (checkBox_user_select_export.Checked == true)
                                {
                                    #region checkBox_user_select_export
                                    if (dt1.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_selected_user.Count; ++i)
                                        {
                                            ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                            for (int j = 0; j < dt1.Rows.Count; ++j)
                                            {
                                                string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                if (layerName1 == lista_selected_user[i])
                                                {
                                                    col_filter_by_layer.Add((ObjectId)dt1.Rows[j][0]);
                                                }
                                            }
                                            string filename = textBox_output_folder.Text;

                                            if (System.IO.Directory.Exists(filename) == false)
                                            {
                                                MessageBox.Show("No valid output folder");
                                                Editor1.SetImpliedSelection(Empty_array);
                                                Editor1.WriteMessage("\nCommand:");
                                                set_enable_true();
                                                return;
                                            }


                                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                                            {
                                                filename = filename + "\\";
                                            }

                                            if (System.IO.Directory.Exists(filename) == true)
                                            {
                                                int incr = 0;
                                                string suff1 = "";
                                                bool exista = true;
                                                do
                                                {

                                                    if (System.IO.File.Exists(filename + lista_selected_user[i] + suff1 + ".shp") == false)
                                                    {
                                                        filename = filename + lista_selected_user[i] + suff1 + ".shp";
                                                        exista = false;
                                                    }
                                                    else
                                                    {

                                                        ++incr;
                                                        suff1 = incr.ToString();
                                                    }

                                                } while (exista == true);


                                            }
                                            ExportSHP("SHP", filename, lista_selected_user[i], true, false, "line", col_filter_by_layer);

                                            for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                                            {
                                                string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                if (layerName1 == lista_selected_user[i])
                                                {
                                                    dt1.Rows[j].Delete();
                                                }
                                            }
                                        }

                                    }

                                    if (dt2.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < lista_selected_user.Count; ++i)
                                        {
                                            ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                            for (int j = 0; j < dt2.Rows.Count; ++j)
                                            {
                                                string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                if (layerName1 == lista_selected_user[i])
                                                {
                                                    col_filter_by_layer.Add((ObjectId)dt2.Rows[j][0]);
                                                }
                                            }
                                            string filename = textBox_output_folder.Text;

                                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                                            {
                                                filename = filename + "\\";
                                            }

                                            if (System.IO.Directory.Exists(filename) == true)
                                            {
                                                int incr = 0;
                                                string suff1 = "";
                                                bool exista = true;
                                                do
                                                {

                                                    if (System.IO.File.Exists(filename + lista_selected_user[i] + suff1 + ".shp") == false)
                                                    {
                                                        filename = filename + lista_selected_user[i] + suff1 + ".shp";
                                                        exista = false;
                                                    }
                                                    else
                                                    {

                                                        ++incr;
                                                        suff1 = incr.ToString();
                                                    }

                                                } while (exista == true);


                                            }
                                            ExportSHP("SHP", filename, lista_selected_user[i], true, false, "point", col_filter_by_layer);

                                            for (int j = dt2.Rows.Count - 1; j >= 0; --j)
                                            {
                                                string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                if (layerName1 == lista_selected_user[i])
                                                {
                                                    dt2.Rows[j].Delete();
                                                }
                                            }
                                        }

                                    }

                                    if (dt3.Rows.Count > 0)
                                    {

                                        for (int i = 0; i < lista_polygon_layers_user.Count; ++i)
                                        {
                                            ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                            for (int j = 0; j < dt3.Rows.Count; ++j)
                                            {
                                                string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                if (layerName1 == lista_polygon_layers_user[i])
                                                {
                                                    col_filter_by_layer.Add((ObjectId)dt3.Rows[j][0]);
                                                }
                                            }

                                            ObjectIdCollection col_polygon = new ObjectIdCollection();
                                            if (col_filter_by_layer.Count > 0)
                                            {
                                                for (int k = 0; k < col_filter_by_layer.Count; ++k)
                                                {
                                                    Polyline polyline2 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as Polyline;
                                                    MPolygon mpolyg1 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as MPolygon;

                                                    if (polyline2 != null)
                                                    {
                                                        MPolygon mpolygon1 = new MPolygon();
                                                        mpolygon1.AppendLoopFromBoundary(polyline2, true, 1e-12);
                                                        mpolygon1.BalanceTree();
                                                        mpolygon1.Elevation = 0;
                                                        mpolygon1.Normal = polyline2.Normal;
                                                        BTrecord.AppendEntity(mpolygon1);
                                                        Trans1.AddNewlyCreatedDBObject(mpolygon1, true);

                                                        copy_od(polyline2.ObjectId, mpolygon1.ObjectId);
                                                        col_polygon.Add(mpolygon1.ObjectId);
                                                    }

                                                    if (mpolyg1 != null)
                                                    {
                                                        col_polygon.Add(mpolyg1.ObjectId);
                                                    }
                                                }
                                            }


                                            string filename = textBox_output_folder.Text;

                                            if (filename.Substring(filename.Length - 1, 1) != "\\")
                                            {
                                                filename = filename + "\\";
                                            }

                                            if (System.IO.Directory.Exists(filename) == true)
                                            {
                                                int incr = 0;
                                                string suff1 = "";
                                                bool exista = true;
                                                do
                                                {

                                                    if (System.IO.File.Exists(filename + lista_polygon_layers_user[i] + suff1 + ".shp") == false)
                                                    {
                                                        filename = filename + lista_polygon_layers_user[i] + suff1 + ".shp";
                                                        exista = false;
                                                    }
                                                    else
                                                    {

                                                        ++incr;
                                                        suff1 = incr.ToString();
                                                    }

                                                } while (exista == true);


                                            }
                                            ExportSHP("SHP", filename, lista_polygon_layers_user[i], true, false, "polygon", col_polygon);



                                            for (int j = dt3.Rows.Count - 1; j >= 0; --j)
                                            {
                                                string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                if (layerName1 == lista_polygon_layers_user[i])
                                                {
                                                    dt3.Rows[j].Delete();
                                                }
                                            }

                                            for (int k = 0; k < col_polygon.Count; ++k)
                                            {
                                                MPolygon mp1 = Trans1.GetObject(col_polygon[k], OpenMode.ForWrite) as MPolygon;
                                                mp1.Erase();
                                            }
                                        }



                                    }

                                    #endregion

                                }
                                else
                                {
                                    #region others

                                    if (linie_in_Polygon_layer == false && open_poly_in_Polygon_layer == false && point_in_line_layer == false)
                                    {
                                        if (dt1.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < lista_polylines_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                                for (int j = 0; j < dt1.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                    if (layerName1 == lista_polylines_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt1.Rows[j][0]);
                                                    }
                                                }
                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_polylines_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_polylines_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_polylines_layers[i], true, false, "line", col_filter_by_layer);

                                                for (int j = dt1.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt1.Rows[j][1]);
                                                    if (layerName1 == lista_polylines_layers[i])
                                                    {
                                                        dt1.Rows[j].Delete();
                                                    }
                                                }
                                            }

                                        }
                                        if (dt2.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < lista_point_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                                for (int j = 0; j < dt2.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                    if (layerName1 == lista_point_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt2.Rows[j][0]);
                                                    }
                                                }
                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_point_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_point_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_point_layers[i], true, false, "point", col_filter_by_layer);

                                                for (int j = dt2.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt2.Rows[j][1]);
                                                    if (layerName1 == lista_point_layers[i])
                                                    {
                                                        dt2.Rows[j].Delete();
                                                    }
                                                }
                                            }

                                        }
                                        if (dt3.Rows.Count > 0)
                                        {
                                           
                                            for (int i = 0; i < lista_polygon_layers.Count; ++i)
                                            {
                                                ObjectIdCollection col_filter_by_layer = new ObjectIdCollection();
                                                for (int j = 0; j < dt3.Rows.Count; ++j)
                                                {
                                                    string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                    if (layerName1 == lista_polygon_layers[i])
                                                    {
                                                        col_filter_by_layer.Add((ObjectId)dt3.Rows[j][0]);
                                                    }
                                                }

                                                ObjectIdCollection col_polygon = new ObjectIdCollection();
                                                if (col_filter_by_layer.Count > 0)
                                                {
                                                    for (int k = 0; k < col_filter_by_layer.Count; ++k)
                                                    {
                                                        Polyline polyline2 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as Polyline;
                                                        MPolygon mpolyg1 = Trans1.GetObject(col_filter_by_layer[k], OpenMode.ForRead) as MPolygon;

                                                        if (polyline2 != null)
                                                        {
                                                            MPolygon mpolygon1 = new MPolygon();
                                                            mpolygon1.AppendLoopFromBoundary(polyline2, true, 1e-12);
                                                            mpolygon1.BalanceTree();
                                                            mpolygon1.Elevation = 0;
                                                            mpolygon1.Normal = polyline2.Normal;
                                                            BTrecord.AppendEntity(mpolygon1);
                                                            Trans1.AddNewlyCreatedDBObject(mpolygon1, true);

                                                            copy_od(polyline2.ObjectId, mpolygon1.ObjectId);
                                                            col_polygon.Add(mpolygon1.ObjectId);
                                                        }

                                                        if (mpolyg1 != null)
                                                        {
                                                            col_polygon.Add(mpolyg1.ObjectId);
                                                        }
                                                    }
                                                }


                                                string filename = textBox_output_folder.Text;

                                                if (filename.Substring(filename.Length - 1, 1) != "\\")
                                                {
                                                    filename = filename + "\\";
                                                }

                                                if (System.IO.Directory.Exists(filename) == true)
                                                {
                                                    int incr = 0;
                                                    string suff1 = "";
                                                    bool exista = true;
                                                    do
                                                    {

                                                        if (System.IO.File.Exists(filename + lista_polygon_layers[i] + suff1 + ".shp") == false)
                                                        {
                                                            filename = filename + lista_polygon_layers[i] + suff1 + ".shp";
                                                            exista = false;
                                                        }
                                                        else
                                                        {

                                                            ++incr;
                                                            suff1 = incr.ToString();
                                                        }

                                                    } while (exista == true);


                                                }
                                                ExportSHP("SHP", filename, lista_polygon_layers[i], true, false, "polygon", col_polygon);



                                                for (int j = dt3.Rows.Count - 1; j >= 0; --j)
                                                {
                                                    string layerName1 = Convert.ToString(dt3.Rows[j][1]);
                                                    if (layerName1 == lista_polygon_layers[i])
                                                    {
                                                        dt3.Rows[j].Delete();
                                                    }
                                                }

                                                for (int k = 0; k < col_polygon.Count; ++k)
                                                {
                                                    MPolygon mp1 = Trans1.GetObject(col_polygon[k], OpenMode.ForWrite) as MPolygon;
                                                    mp1.Erase();
                                                }
                                            }



                                        }
                                    }



                                    if (point_in_line_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have lines into the point layers or points inside lines layers\r\nsee colorindex 42");
                                    }


                                    if (open_poly_in_Polygon_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have at least one open polyline that you want to export it as a polygon\r\nsee colorindex 42");

                                    }

                                    if (linie_in_Polygon_layer == true)
                                    {
                                        MessageBox.Show("Operation aborted!\r\nyou have at least one item that is not a polyline that you want to export it as a polygon\r\nsee colorindex 42");

                                    }
                                    #endregion
                                }





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

            }

        }


        public void ExportSHP(string format, string filename, string layername, bool isODTable, bool isLinkTemplate, string geomtryType, ObjectIdCollection object_id_col)
        {
            Exporter exporter = null;
            try
            {
                MapApplication mapApp = HostMapApplicationServices.Application;
                exporter = mapApp.Exporter;
                // Initiate the exporter
                exporter.Init(format, filename);
                if (object_id_col != null && object_id_col.Count > 0)
                {
                    exporter.ExportAll = false;

                    exporter.SetSelectionSet(object_id_col);
                }
                else
                {
                    //exporter.ExportAll = true;
                }

                GeometryType geomtryType1 = (GeometryType)Enum.Parse(typeof(GeometryType), geomtryType, true);
                exporter.SetStorageOptions(StorageType.FileOneEntityType, geomtryType1, string.Empty);

                // Get Data mapping object
                ExpressionTargetCollection dataMapping = null;
                dataMapping = exporter.GetExportDataMappings();
                dataMapping.Clear();

                // Set ObjectData data mapping if isODTable is true
                if (isODTable == true && MapODData(dataMapping, layername) == true)
                {
                    // Reset Data mapping with Object data and Link template keys		
                    exporter.SetExportDataMappings(dataMapping);
                }
                // If layerFilter isn't null, set the layer filter to export layer by layer
                if (null != layername)
                {
                    exporter.LayerFilter = layername;
                }

                // Do the exporting and log the result
                ExportResults results;
                results = exporter.Export(true);

                Utility.ShowMsg("\nExporting succeeded.");
            }
            catch (MapException e)
            {

            }
            finally
            {

            }
        }


        public bool MapODData(ExpressionTargetCollection mapping, string tablename)
        {
            MapApplication mapApi = HostMapApplicationServices.Application;
            ProjectModel proj = mapApi.ActiveProject;
            Tables tables = proj.ODTables;
            if (tables.IsTableDefined(tablename) == true)
            {
                try
                {
                    Autodesk.Gis.Map.ObjectData.Table table = tables[tablename];
                    FieldDefinitions definitions = table.FieldDefinitions;
                    for (int j = 0; j < definitions.Count; j++)
                    {
                        FieldDefinition column = null;
                        column = definitions[j];
                        // fieldName is the OD table field name in the data mapping. It should be 
                        // in the format:fieldName&tableName. 
                        // newFieldName is the attribute field name of exported-to file
                        string ODfieldName = ":" + column.Name + "@" + tablename;
                        string shpFieldName = column.Name;

                        mapping.Add(ODfieldName, shpFieldName);
                    }
                }
                catch (MapImportExportException)
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }

        }
        private List<string> get_layers_from_dwg()
        {
            List<string> lista_layers = new List<string>();

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument();
            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
            {
                Autodesk.AutoCAD.DatabaseServices.LayerTable layer_table = (Autodesk.AutoCAD.DatabaseServices.LayerTable)Trans1.GetObject(ThisDrawing.Database.LayerTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                System.Data.DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("ln", typeof(string));
                foreach (ObjectId Layer_id in layer_table)
                {
                    LayerTableRecord Layer1 = (LayerTableRecord)Trans1.GetObject(Layer_id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    string Name_of_layer = Layer1.Name;
                    if (Name_of_layer.Contains("|") == false && Name_of_layer.Contains("$") == false && Layer1.IsFrozen == false && Layer1.IsOff == false)
                    {
                        dt1.Rows.Add();
                        dt1.Rows[dt1.Rows.Count - 1][0] = Name_of_layer;
                    }
                }
                System.Data.DataTable dt2 = Functions.Sort_data_table(dt1, "ln");
                for (int i = 0; i < dt2.Rows.Count; ++i)
                {
                    lista_layers.Add(dt2.Rows[i][0].ToString());
                }
                Trans1.Commit();
            }
            return lista_layers;
        }



        private void button_multiselect_Click(object sender, EventArgs e)
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
                        ObjectId[] objid1;
                        objid1 = new ObjectId[0];
                        foreach (DataGridViewCell cell1 in DataGridView_data.SelectedCells)
                        {
                            if (cell1.ColumnIndex == 0)
                            {
                                Array.Resize(ref objid1, objid1.Length + 1);
                                objid1[objid1.Length - 1] = (ObjectId)cell1.Value;
                            }
                        }
                        Editor1.SetImpliedSelection(objid1);
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        private void button_select_centerline_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            Data_table_centerline = Functions.Creaza_centerline_datatable_structure();
                            Data_table_centerline.Columns.Add("Bulge", typeof(double));
                            Set_centerline_label_to_red();
                            delete_station_labels();
                            Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_optionsCL = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect Centerline:");
                            Prompt_optionsCL.SetRejectMessage("\nYou did not selected a polyline (2d or 3d)");
                            Prompt_optionsCL.AddAllowedClass(typeof(Polyline), true);
                            Prompt_optionsCL.AddAllowedClass(typeof(Polyline3d), true);
                            this.WindowState = FormWindowState.Minimized;

                            PromptEntityResult Rezultat_CL = Editor1.GetEntity(Prompt_optionsCL);
                            if (Rezultat_CL.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }

                            Curve Curba1 = Trans1.GetObject(Rezultat_CL.ObjectId, OpenMode.ForRead) as Curve;
                            if (Curba1 == null)
                            {
                                MessageBox.Show("you did not select a polyline or a polyline3d");
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }

                            Polyline Poly1 = null;
                            Polyline3d Poly3 = null;
                            if (Curba1 is Polyline)
                            {
                                Poly1 = (Polyline)Curba1;
                                Poly3 = null;
                                project_type = "2d";
                            }
                            else if (Curba1 is Polyline3d)
                            {
                                Poly3 = (Polyline3d)Curba1;
                                Poly1 = Functions.Build_2dpoly_from_3d(Poly3);
                                project_type = "3d";
                            }
                            else
                            {
                                MessageBox.Show("you did not select a polyline or a polyline3d");
                                Freeze_operations = false;
                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }


                            if (checkBox_reverse_direction.Checked == false)
                            {
                                for (int i = 0; i < Poly1.NumberOfVertices; ++i)
                                {
                                    double x2 = Poly1.GetPointAtParameter(i).X;
                                    double y2 = Poly1.GetPointAtParameter(i).Y;
                                    double z2 = Poly1.GetPointAtParameter(i).Z;
                                    if (Poly3 != null)
                                    {
                                        z2 = Poly3.GetPointAtParameter(i).Z;
                                    }
                                    Data_table_centerline.Rows.Add();
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_x] = x2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_y] = y2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_z] = z2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = Poly1.GetBulgeAt(i);
                                }
                            }
                            else
                            {
                                for (int i = Poly1.NumberOfVertices - 1; i >= 0; --i)
                                {
                                    double x2 = Poly1.GetPointAtParameter(i).X;
                                    double y2 = Poly1.GetPointAtParameter(i).Y;
                                    double z2 = Poly1.GetPointAtParameter(i).Z;
                                    if (Poly3 != null)
                                    {
                                        z2 = Poly3.GetPointAtParameter(i).Z;
                                    }
                                    Data_table_centerline.Rows.Add();
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_x] = x2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_y] = y2;
                                    Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1][Col_z] = z2;
                                    if (i - 1 >= 0)
                                    {
                                        Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = -Poly1.GetBulgeAt(i - 1);
                                    }
                                    else
                                    {
                                        Data_table_centerline.Rows[Data_table_centerline.Rows.Count - 1]["Bulge"] = 0;
                                    }
                                }
                            }
                            Trans1.Commit();
                            Set_centerline_label_to_green(Curba1.ObjectId.Handle.Value.ToString());
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            Freeze_operations = false;
            this.WindowState = FormWindowState.Normal;
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
            string Curent_system = Acmap.GetMapSRS();
            if (Curent_system != "")
            {
                OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                set_label_cs_to_green(CoordSys1.CsCode);
            }
            else
            {
                set_label_cs_to_red();
            }
        }


        public void delete_station_labels()
        {
            if (col_station_labels.Count > 0)
            {
                using (ObjectIdCollection col1 = new ObjectIdCollection())
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            foreach (string handle1 in col_station_labels)
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    try
                                    {
                                        Entity ent1 = Trans1.GetObject(id1, OpenMode.ForRead) as Entity;
                                        if (ent1 != null)
                                        {

                                            ent1.UpgradeOpen();
                                            ent1.Erase();
                                            col1.Add(id1);


                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }

                    if (col1.Count > 0)
                    {
                        foreach (ObjectId id1 in col1)
                        {
                            string handle1 = id1.Handle.Value.ToString();
                            if (col_station_labels.Contains(handle1) == true)
                            {
                                col_station_labels.Remove(handle1);
                            }
                        }
                    }
                }
            }
        }

        private void Set_centerline_label_to_green(string handle1)
        {
            label_cl_loaded.Text = "CL loaded \r\nhandle# - " + handle1;
            label_cl_loaded.ForeColor = Color.LimeGreen;

            if (project_type == "2d")
            {
                label_sta.Text = "2D Station:";
                label_mp.Text = "MP(2D):";
            }
            else
            {
                label_sta.Text = "3D Station:";
                label_mp.Text = "MP(3D):";
            }

            textBox_offset.Text = "";
            textBox_station.Text = "";
            textBox_mp.Text = "";
            textBox_zoom_to.Text = "";
            textBox_x.Text = "";
            textBox_y.Text = "";
            textBox_z.Text = "";
            textBox_lat.Text = "";
            textBox_long.Text = "";

        }

        public void set_label_cs_to_red()
        {
            label_cs.Text = "NO coordinate system set";
            label_cs.ForeColor = Color.Red;
        }

        public void set_label_cs_to_green(string cs_name)
        {
            label_cs.Text = cs_name;
            label_cs.ForeColor = Color.LimeGreen;
        }
        private void Set_centerline_label_to_red()
        {
            label_cl_loaded.Text = "CL not loaded";
            label_cl_loaded.ForeColor = Color.Red;
            label_sta.Text = "2D Station:";
            label_mp.Text = "MP(2D):";
            textBox_offset.Text = "";
            textBox_station.Text = "";
            textBox_mp.Text = "";
            textBox_zoom_to.Text = "";
            textBox_x.Text = "";
            textBox_y.Text = "";
            textBox_z.Text = "";
            textBox_lat.Text = "";
            textBox_long.Text = "";
        }

        private void button_redraw_centerline_Click(object sender, EventArgs e)
        {
            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_centerline != null)
                {
                    if (Data_table_centerline.Rows.Count > 0)
                    {
                        if (checkBox_reverse_direction.Checked == true)
                        {
                            Data_table_centerline = reverse_cl();
                        }

                        try
                        {
                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                Functions.Creaza_layer(layer_no_plot, 30, false);
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                    if (project_type == "2d")
                                    {
                                        Polyline new_poly2d = new Polyline();
                                        for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                                        {
                                            double x = 0;
                                            double y = 0;
                                            double bulge1 = 0;

                                            if (Data_table_centerline.Rows[i][Col_x] != DBNull.Value)
                                            {
                                                x = (double)Data_table_centerline.Rows[i][Col_x];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no X value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][Col_y] != DBNull.Value)
                                            {
                                                y = (double)Data_table_centerline.Rows[i][Col_y];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Columns.Contains("Bulge") == true)
                                            {
                                                if (Data_table_centerline.Rows[i]["Bulge"] != DBNull.Value)
                                                {
                                                    bulge1 = Convert.ToDouble(Data_table_centerline.Rows[i]["Bulge"]);
                                                }
                                            }
                                            new_poly2d.AddVertexAt(i, new Point2d(x, y), bulge1, 0, 0);

                                        }

                                        new_poly2d.Layer = layer_no_plot;
                                        BTrecord.AppendEntity(new_poly2d);
                                        Trans1.AddNewlyCreatedDBObject(new_poly2d, true);

                                        cl_id_for_temp = new_poly2d.ObjectId.Handle.Value.ToString();

                                    }
                                    else
                                    {
                                        Polyline3d new_poly3d = new Polyline3d();
                                        new_poly3d.SetDatabaseDefaults();
                                        new_poly3d.Layer = layer_no_plot;
                                        BTrecord.AppendEntity(new_poly3d);
                                        Trans1.AddNewlyCreatedDBObject(new_poly3d, true);

                                        for (int i = 0; i < Data_table_centerline.Rows.Count; ++i)
                                        {
                                            double x = 0;
                                            double y = 0;
                                            double z = 0;

                                            if (Data_table_centerline.Rows[i][Col_x] != DBNull.Value)
                                            {
                                                x = (double)Data_table_centerline.Rows[i][Col_x];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no X value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][Col_y] != DBNull.Value)
                                            {
                                                y = (double)Data_table_centerline.Rows[i][Col_y];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }
                                            if (Data_table_centerline.Rows[i][Col_z] != DBNull.Value)
                                            {
                                                z = (double)Data_table_centerline.Rows[i][Col_z];
                                            }
                                            else
                                            {

                                                Freeze_operations = false;
                                                MessageBox.Show("no Y value for centerline in row " + (i).ToString());
                                                return;
                                            }

                                            PolylineVertex3d Vertex_new = new PolylineVertex3d(new Point3d(x, y, z));
                                            new_poly3d.AppendVertex(Vertex_new);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new, true);

                                        }

                                        cl_id_for_temp = new_poly3d.ObjectId.Handle.Value.ToString();

                                    }
                                    Trans1.Commit();

                                    if (checkBox_reverse_direction.Checked == true)
                                    {
                                        Data_table_centerline = reverse_cl();
                                    }



                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    Freeze_operations = false;
                }
            }
        }
        private System.Data.DataTable reverse_cl()
        {
            System.Data.DataTable dt1 = Functions.Creaza_centerline_datatable_structure();
            dt1.Columns.Add("Bulge", typeof(double));

            if (Data_table_centerline != null)
            {
                if (Data_table_centerline.Rows.Count > 0)
                {
                    for (int i = Data_table_centerline.Rows.Count - 1; i >= 0; --i)
                    {
                        dt1.Rows.Add();
                        for (int j = 0; j < Data_table_centerline.Columns.Count; ++j)
                        {
                            dt1.Rows[dt1.Rows.Count - 1][j] = Data_table_centerline.Rows[i][j];
                        }
                    }

                }
            }

            return dt1;
        }

        private void button_create_station_label_Click(object sender, EventArgs e)
        {

            double spacing_major = 500;
            double spacing_minor = 100;
            double tick_major = 20;
            double tick_minor = 10;
            double texth = 4;
            double gap1 = 4;
            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }

            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);

            if (Freeze_operations == false)
            {
                Freeze_operations = true;

                if (Data_table_centerline != null)
                {
                    if (Data_table_centerline.Rows.Count > 0)
                    {
                        try
                        {
                            delete_station_labels();

                            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                            using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                            {
                                Functions.Creaza_layer(layer_no_plot, 30, false);
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                                    Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                                    Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                                    Polyline3d Poly3D = null;

                                    if (project_type == "3d") Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                                    if (project_type == "2d")
                                    {
                                        if (Poly2D.Length >= spacing_major)
                                        {
                                            int no_major = Convert.ToInt32(Math.Floor((Poly2D.Length) / spacing_major)) + 2;



                                            double first_label_major = spacing_major * Math.Ceiling(start1 / spacing_major);
                                            double len_stationed_major = Poly2D.Length - (first_label_major - start1);


                                            do
                                            {
                                                if (no_major * spacing_major >= len_stationed_major)
                                                {
                                                    no_major = no_major - 1;
                                                }
                                            } while (no_major * spacing_major >= len_stationed_major);



                                            if (no_major > 0)
                                            {
                                                for (int i = 0; i <= no_major; ++i)
                                                {
                                                    Point3d pt0 = Poly2D.GetPointAtDist((first_label_major - start1) + i * spacing_major);
                                                    double label_major = first_label_major + i * spacing_major;
                                                    Autodesk.AutoCAD.DatabaseServices.Line Big1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;



                                                    Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    Big1.Layer = layer_no_plot;
                                                    Big1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);

                                                    col_station_labels.Add(Big1.ObjectId.Handle.Value.ToString());

                                                    Autodesk.AutoCAD.DatabaseServices.Line l_t = new Autodesk.AutoCAD.DatabaseServices.Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1);

                                                    mt1.Layer = layer_no_plot;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    col_station_labels.Add(mt1.ObjectId.Handle.Value.ToString());
                                                }
                                            }
                                        }

                                        if (Poly2D.Length >= spacing_minor)
                                        {

                                            int no_minor = Convert.ToInt32(Math.Floor((Poly2D.Length) / spacing_minor)) + 2;




                                            double first_label_minor = spacing_minor * Math.Ceiling(start1 / spacing_minor);
                                            double len_stationed_minor = Poly2D.Length - (first_label_minor - start1);



                                            do
                                            {
                                                if (no_minor * spacing_minor >= len_stationed_minor)
                                                {
                                                    no_minor = no_minor - 1;
                                                }
                                            } while (no_minor * spacing_minor >= len_stationed_minor);


                                            if (no_minor > 0)
                                            {
                                                for (int i = 0; i <= no_minor; ++i)
                                                {
                                                    Point3d pt0 = Poly2D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                                    double label_major = first_label_minor + i * spacing_minor;
                                                    Autodesk.AutoCAD.DatabaseServices.Line small1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly2D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly2D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;
                                                    small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));
                                                    small1.Layer = layer_no_plot;
                                                    small1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(small1);
                                                    Trans1.AddNewlyCreatedDBObject(small1, true);
                                                    col_station_labels.Add(small1.ObjectId.Handle.Value.ToString());
                                                }
                                            }
                                        }
                                    }
                                    if (project_type == "3d")
                                    {
                                        if (Poly3D.Length >= spacing_major)
                                        {
                                            int no_major = Convert.ToInt32(Math.Floor((Poly3D.Length) / spacing_major)) + 2;


                                            double first_label_major = spacing_major * Math.Ceiling(start1 / spacing_major);
                                            double len_stationed_major = Poly3D.Length - (first_label_major - start1);



                                            do
                                            {
                                                if (no_major * spacing_major >= len_stationed_major)
                                                {
                                                    no_major = no_major - 1;
                                                }
                                            } while (no_major * spacing_major >= len_stationed_major);




                                            if (no_major > 0)
                                            {
                                                for (int i = 0; i <= no_major; ++i)
                                                {
                                                    Point3d pt0 = Poly3D.GetPointAtDist((first_label_major - start1) + i * spacing_major);


                                                    double label_major = first_label_major + i * spacing_major;
                                                    Autodesk.AutoCAD.DatabaseServices.Line Big1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_major / 2, pt0.Y, 0), new Point3d(pt0.X + tick_major / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;

                                                    Big1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    Big1.Layer = layer_no_plot;
                                                    Big1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(Big1);
                                                    Trans1.AddNewlyCreatedDBObject(Big1, true);

                                                    col_station_labels.Add(Big1.ObjectId.Handle.Value.ToString());

                                                    Autodesk.AutoCAD.DatabaseServices.Line l_t = new Autodesk.AutoCAD.DatabaseServices.Line(Big1.StartPoint, Big1.EndPoint);
                                                    l_t.TransformBy(Matrix3d.Scaling((Big1.Length + gap1) / Big1.Length, Big1.EndPoint));

                                                    MText mt1 = creaza_mtext_sta(l_t.StartPoint, Functions.Get_chainage_from_double(first_label_major + i * spacing_major, "f", 0), texth, bear1);

                                                    mt1.Layer = layer_no_plot;
                                                    BTrecord.AppendEntity(mt1);
                                                    Trans1.AddNewlyCreatedDBObject(mt1, true);
                                                    col_station_labels.Add(mt1.ObjectId.Handle.Value.ToString());

                                                }
                                            }
                                        }

                                        if (Poly3D.Length >= spacing_minor)
                                        {
                                            int no_minor = Convert.ToInt32(Math.Floor((Poly3D.Length) / spacing_minor)) + 2;


                                            double first_label_minor = spacing_minor * Math.Ceiling(start1 / spacing_minor);
                                            double len_stationed_minor = Poly3D.Length - (first_label_minor - start1);

                                            do
                                            {
                                                if (no_minor * spacing_minor >= len_stationed_minor)
                                                {
                                                    no_minor = no_minor - 1;
                                                }
                                            } while (no_minor * spacing_minor >= len_stationed_minor);


                                            if (no_minor > 0)
                                            {
                                                for (int i = 0; i <= no_minor; ++i)
                                                {
                                                    Point3d pt0 = Poly3D.GetPointAtDist((first_label_minor - start1) + i * spacing_minor);
                                                    double label_major = first_label_minor + i * spacing_minor;
                                                    Autodesk.AutoCAD.DatabaseServices.Line small1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(pt0.X - tick_minor / 2, pt0.Y, 0), new Point3d(pt0.X + tick_minor / 2, pt0.Y, 0));

                                                    double param1 = Poly2D.GetParameterAtPoint(Poly2D.GetClosestPointTo(pt0, Vector3d.ZAxis, false));
                                                    double param2 = param1 + 1;
                                                    if (Poly2D.EndParam < param2)
                                                    {
                                                        param1 = Poly2D.EndParam - 1;
                                                        param2 = Poly2D.EndParam;
                                                    }


                                                    Point3d point1 = Poly3D.GetPointAtParameter(Math.Floor(param1));

                                                    Point3d point2 = Poly3D.GetPointAtParameter(Math.Floor(param2));

                                                    double bear1 = Functions.GET_Bearing_rad(point1.X, point1.Y, point2.X, point2.Y);

                                                    double rot1 = bear1 - Math.PI / 2;

                                                    small1.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt0));

                                                    small1.Layer = layer_no_plot;
                                                    small1.ColorIndex = 256;
                                                    BTrecord.AppendEntity(small1);
                                                    Trans1.AddNewlyCreatedDBObject(small1, true);

                                                    col_station_labels.Add(small1.ObjectId.Handle.Value.ToString());
                                                }
                                            }
                                        }

                                    }

                                    if (project_type == "3d") delete_cl(Poly3D);
                                    delete_cl(Poly2D);
                                    Trans1.Commit();
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }

                    Freeze_operations = false;
                }
            }
        }

        public static MText creaza_mtext_sta(Point3d pt_ins, string continut, double texth, double rot1)
        {


            MText mtext1 = new MText();
            mtext1.Attachment = AttachmentPoint.BottomCenter;
            mtext1.Contents = continut;
            mtext1.TextHeight = texth;
            mtext1.BackgroundFill = true;
            mtext1.UseBackgroundColor = true;
            mtext1.BackgroundScaleFactor = 1.2;
            mtext1.Location = pt_ins;
            mtext1.Rotation = rot1;
            mtext1.ColorIndex = 256;


            return mtext1;


        }


        public void delete_zoom_labels()
        {
            if (col_labels_zoom.Count > 0)
            {
                using (ObjectIdCollection col1 = new ObjectIdCollection())
                {

                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            foreach (string handle1 in col_labels_zoom)
                            {
                                ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                                if (id1 != ObjectId.Null)
                                {
                                    try
                                    {
                                        MText mt_zoom = Trans1.GetObject(id1, OpenMode.ForRead) as MText;
                                        if (mt_zoom != null)
                                        {
                                            if (mt_zoom.Layer == layer_no_plot)
                                            {
                                                if (mt_zoom.TextHeight == 0.1)
                                                {
                                                    mt_zoom.UpgradeOpen();
                                                    mt_zoom.Erase();
                                                    col1.Add(id1);
                                                }
                                            }
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                    }
                                }
                            }
                            Trans1.Commit();
                        }
                    }

                    if (col1.Count > 0)
                    {
                        foreach (ObjectId id1 in col1)
                        {
                            string handle1 = id1.Handle.Value.ToString();
                            if (col_labels_zoom.Contains(handle1) == true)
                            {
                                col_labels_zoom.Remove(handle1);
                            }
                        }
                    }
                }
            }
        }




        public Point3d Convert_point_to_new_CS(Point3d Point1, string to_coord_system)
        {
            Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();

            string Curent_system = Acmap.GetMapSRS();
            if (Curent_system == "")
            {
                set_label_cs_to_red();
                return new Point3d();
            }

            Point3d Point2 = new Point3d();
            OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
            OSGeo.MapGuide.MgCoordinateSystemCatalog Catalog1 = Coord_factory1.GetCatalog();
            OSGeo.MapGuide.MgCoordinateSystemDictionary Dictionary1 = Catalog1.GetCoordinateSystemDictionary();
            OSGeo.MapGuide.MgCoordinateSystemEnum Enum1 = Dictionary1.GetEnum();

            OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);

            OSGeo.MapGuide.MgCoordinateSystem CoordSys2 = Dictionary1.GetCoordinateSystem(to_coord_system);

            OSGeo.MapGuide.MgCoordinateSystemTransform Transform1 = Coord_factory1.GetTransform(CoordSys1, CoordSys2);
            OSGeo.MapGuide.MgCoordinate Coord1 = Transform1.Transform(Point1.X, Point1.Y);

            Point2 = new Point3d(Coord1.X, Coord1.Y, 0);

            set_label_cs_to_green(CoordSys1.CsCode);

            return Point2;
        }

        private void delete_cl(Entity Poly1)
        {
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock lock1 = ThisDrawing.LockDocument())
            {
                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                {

                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                    Entity ent1 = Trans1.GetObject(Poly1.ObjectId, OpenMode.ForWrite) as Entity;
                    if (ent1 != null)
                    {
                        ent1.Erase();
                        Trans1.Commit();
                    }
                }
            }
        }
        public void delete_cl_from_redraw(string handle1)
        {
            if (handle1 != null)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        ObjectId id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);
                        if (id1 != ObjectId.Null)
                        {
                            try
                            {
                                Entity ent1 = Trans1.GetObject(id1, OpenMode.ForWrite) as Entity;
                                if (ent1 != null)
                                {
                                    ent1.Erase();
                                    Trans1.Commit();
                                }
                            }
                            catch (System.Exception ex)
                            {
                            }
                        }

                    }

                }
            }
        }
        private Polyline creeaza_poly2d(System.Data.DataTable dt_cl)
        {
            Polyline Poly2D = new Polyline();

            if (dt_cl.Rows.Count > 0)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;



                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


                        int index1 = 0;

                        for (int i = 0; i < dt_cl.Rows.Count; ++i)
                        {
                            double x = 0;
                            double y = 0;

                            if (dt_cl.Rows[i][Col_x] != DBNull.Value)
                            {
                                x = (double)dt_cl.Rows[i][Col_x];
                                if (dt_cl.Rows[i][Col_y] != DBNull.Value)
                                {
                                    y = (double)dt_cl.Rows[i][Col_y];

                                    double bulge1 = 0;
                                    if (dt_cl.Columns.Contains("Bulge") == true)
                                    {
                                        bulge1 = Convert.ToDouble(dt_cl.Rows[i]["Bulge"]);
                                    }


                                    Poly2D.AddVertexAt(index1, new Point2d(x, y), bulge1, 0, 0);
                                    Poly2D.Elevation = 0;

                                    index1 = index1 + 1;
                                }
                            }
                        }

                        BTrecord.AppendEntity(Poly2D);
                        Trans1.AddNewlyCreatedDBObject(Poly2D, true);

                        Trans1.Commit();

                    }
                }

            }


            return Poly2D;


        }

        private void checkBox_reverse_direction_CheckedChanged(object sender, EventArgs e)
        {
            Data_table_centerline = reverse_cl();

            button_create_station_label_Click(sender, e);

        }

        private void button_zoom_to_Click(object sender, EventArgs e)
        {
            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }


            string station_ammount = textBox_zoom_to.Text;
            if (Functions.IsNumeric(textBox_zoom_to.Text.Replace("+", "")) == false)
            {
                MessageBox.Show("station is not specified properly");
                return;
            }

            delete_zoom_labels();


            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);


            double Sta_pt = Convert.ToDouble(station_ammount.Replace("+", ""));


            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(layer_no_plot, 30, false);

                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                            Polyline3d Poly3D = null;

                            if (project_type == "3d")

                            {
                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                            }



                            string continut = "";

                            double Station1 = -123.123;

                            if (comboBox_label_type.Text == "by STA")
                            {
                                Station1 = Sta_pt;

                                if (project_type == "2d")
                                {
                                    if (Poly2D.Length < Sta_pt - start1)
                                    {
                                        MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly2D.GetPointAtDist(Sta_pt - start1);
                                    continut = "STA2d=" + Functions.Get_chainage_from_double(Sta_pt, "f", 2) + "\r\nMP2d=" + Functions.Get_String_Rounded(Sta_pt / 5280, 2);

                                }
                                else
                                {
                                    if (Poly3D.Length < Sta_pt - start1)
                                    {
                                        MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }

                                    pt_on_poly = Poly3D.GetPointAtDist(Sta_pt - start1);
                                    continut = "STA3d=" + Functions.Get_chainage_from_double(Sta_pt, "f", 2) + "\r\nMP3d=" + Functions.Get_String_Rounded(Sta_pt / 5280, 2);
                                }
                            }
                            else
                            {
                                Station1 = Sta_pt * 5280;

                                if (project_type == "2d")
                                {
                                    if (Poly2D.Length < Sta_pt * 5280 - start1)
                                    {
                                        MessageBox.Show("the mp you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly2D.GetPointAtDist(Sta_pt * 5280 - start1);
                                    continut = "STA2d=" + Functions.Get_chainage_from_double(Sta_pt * 5280, "f", 2) + "\r\nMP2d=" + Functions.Get_String_Rounded(Sta_pt, 2);
                                }
                                else
                                {
                                    if (Poly3D.Length < Sta_pt * 5280 - start1)
                                    {
                                        MessageBox.Show("the mp you specified minus the starting station is larger than polyline length ");
                                        Freeze_operations = false;
                                        return;
                                    }
                                    pt_on_poly = Poly3D.GetPointAtDist(Sta_pt * 5280 - start1);
                                    continut = "STA3d=" + Functions.Get_chainage_from_double(Sta_pt * 5280, "f", 2) + "\r\nMP3d=" + Functions.Get_String_Rounded(Sta_pt, 2);
                                }
                            }

                            textBox_station.Text = Functions.Get_chainage_from_double(Station1, "f", 2);
                            textBox_mp.Text = Functions.Get_String_Rounded(Station1 / 5280, 2);
                            textBox_offset.Text = "0";
                            textBox_zoom_to.Text = "";


                            textBox_x.Text = Functions.Get_String_Rounded(pt_on_poly.X, 4);
                            textBox_y.Text = Functions.Get_String_Rounded(pt_on_poly.Y, 4);
                            textBox_z.Text = Functions.Get_String_Rounded(pt_on_poly.Z, 4);
                            Point3d LL = Convert_point_to_new_CS(pt_on_poly, "LL84");

                            textBox_long.Text = Functions.Get_DMS(LL.X, 2);
                            textBox_lat.Text = Functions.Get_DMS(LL.Y, 2);



                            MText mt1 = Functions.creaza_mtext_label(pt_on_poly, continut, 0.1);
                            mt1.Layer = layer_no_plot;
                            BTrecord.AppendEntity(mt1);
                            Trans1.AddNewlyCreatedDBObject(mt1, true);
                            col_labels_zoom.Add(mt1.ObjectId.Handle.Value.ToString());

                            zoom_to_Point(pt_on_poly, 40);
                            picked_pt = pt_on_poly;
                            if (project_type == "3d") delete_cl(Poly3D);
                            delete_cl(Poly2D);
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();

                string Curent_system = Acmap.GetMapSRS();
                if (Curent_system != "")
                {
                    OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                    OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                    set_label_cs_to_green(CoordSys1.CsCode);
                }
                else
                {
                    set_label_cs_to_red();
                }


                Editor1.SetImpliedSelection(Empty_array);
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }
        }
        private void zoom_to_Point(Point3d pt, double zoom_delta_distance)
        {

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;


            try
            {
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        try
                        {



                            Point3d minx = new Point3d(pt.X - zoom_delta_distance, pt.Y - zoom_delta_distance, 0);
                            Point3d maxx = new Point3d(pt.X + zoom_delta_distance, pt.Y + zoom_delta_distance, 0);

                            using (Autodesk.AutoCAD.GraphicsSystem.Manager GraphicsManager = ThisDrawing.GraphicsManager)
                            {

                                int Cvport = Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

                                //from here 2015 dlls:
                                Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor kd = new Autodesk.AutoCAD.GraphicsSystem.KernelDescriptor();
                                kd.addRequirement(Autodesk.AutoCAD.UniqueString.Intern("3D Drawing"));
                                Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.ObtainAcGsView(Cvport, kd);
                                // to here 2015 dlls

                                //from here 2013 dlls:

                                //Autodesk.AutoCAD.GraphicsSystem.View view = GraphicsManager.GetGsView(Cvport, true);

                                // to here 2013 dlls

                                if (view != null)
                                {
                                    using (view)
                                    {

                                        view.ZoomExtents(minx, maxx);

                                        view.Zoom(0.95);//<--optional 
                                        GraphicsManager.SetViewportFromView(Cvport, view, true, true, false);

                                    }
                                }
                                Trans1.Commit();
                            }


                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
            }







            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        private void button_station_at_point_Click(object sender, EventArgs e)
        {

            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                picked_pt = new Point3d(123.123, 123.123, 123.123);
                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                return;
            }

            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);
            delete_zoom_labels();

            if (Freeze_operations == false)
            {
                ObjectId[] Empty_array = null;
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            Polyline Poly2D = null;
                            Polyline3d Poly3D = null;
                            if (project_type == "2d")
                            {
                                Poly2D = creeaza_poly2d(Data_table_centerline);

                            }
                            else
                            {
                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);
                                Poly2D = creeaza_poly2d(Data_table_centerline);
                            }

                            this.WindowState = FormWindowState.Minimized;

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                            PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify point:");
                            PP1.AllowNone = true;
                            Point_res1 = Editor1.GetPoint(PP1);

                            textBox_offset.Text = "";
                            textBox_station.Text = "";
                            textBox_mp.Text = "";
                            textBox_zoom_to.Text = "";
                            textBox_x.Text = "";
                            textBox_y.Text = "";
                            textBox_z.Text = "";
                            textBox_lat.Text = "";
                            textBox_long.Text = "";

                            if (Point_res1.Status != PromptStatus.OK)
                            {
                                Freeze_operations = false;
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                picked_pt = new Point3d(123.123, 123.123, 123.123);
                                pt_on_poly = new Point3d(123.123, 123.123, 123.123);
                                this.WindowState = FormWindowState.Normal;
                                return;
                            }



                            picked_pt = Point_res1.Value.TransformBy(curent_ucs_matrix);
                            pt_on_poly = Poly2D.GetClosestPointTo(picked_pt, Vector3d.ZAxis, false);
                            double Station1 = Math.Round(start1 + Poly2D.GetDistAtPoint(pt_on_poly), 2);
                            delete_cl(Poly2D);


                            if (project_type == "3d")
                            {
                                double param1 = Poly2D.GetParameterAtPoint(pt_on_poly);
                                Station1 = Math.Round(start1 + Poly3D.GetDistanceAtParameter(param1), 2);
                                delete_cl(Poly3D);
                            }

                            textBox_station.Text = Functions.Get_chainage_from_double(Station1, "f", 2);
                            textBox_mp.Text = Functions.Get_String_Rounded(Station1 / 5280, 2);
                            textBox_offset.Text = Functions.Get_String_Rounded(Math.Round(new Point3d(picked_pt.X, picked_pt.Y, 0).DistanceTo(new Point3d(pt_on_poly.X, pt_on_poly.Y, 0)), 2), 2);
                            textBox_x.Text = Functions.Get_String_Rounded(picked_pt.X, 4);
                            textBox_y.Text = Functions.Get_String_Rounded(picked_pt.Y, 4);
                            textBox_z.Text = Functions.Get_String_Rounded(picked_pt.Z, 4);
                            Point3d LL = Convert_point_to_new_CS(picked_pt, "LL84");

                            textBox_long.Text = Functions.Get_DMS(LL.X, 2);
                            textBox_lat.Text = Functions.Get_DMS(LL.Y, 2);

                            if (checkBox_always_create_label.Checked == true)
                            {
                                Freeze_operations = false;
                                button_create_label_Click(sender, e);
                            }


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
                Freeze_operations = false;
                this.WindowState = FormWindowState.Normal;
            }


        }

        private void button_create_label_Click(object sender, EventArgs e)
        {
            if (Data_table_centerline == null)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }
            if (Data_table_centerline.Rows.Count == 0)
            {
                MessageBox.Show("you did not selected any centerline");
                return;
            }

            string start_ammount = textBox_start_station.Text;

            if (Functions.IsNumeric(start_ammount.Replace("+", "")) == false)
            {
                MessageBox.Show("the start station is not specified");
                return;
            }


            if (textBox_station.Text == "")
            {
                MessageBox.Show("First you have to specify a point");
                return;
            }



            string sta_ammount = textBox_station.Text;
            string mp_ammount = textBox_mp.Text;
            double sta1 = Convert.ToDouble(sta_ammount.Replace("+", ""));
            double mp1 = Convert.ToDouble(mp_ammount);
            double start1 = Math.Round(Convert.ToDouble(start_ammount.Replace("+", "")), 2);


            if (Freeze_operations == false)
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                HostApplicationServices.WorkingDatabase = ThisDrawing.Database;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                try
                {
                    Freeze_operations = true;
                    using (DocumentLock lock1 = ThisDrawing.LockDocument())
                    {
                        Functions.Creaza_layer(layer_no_plot, 30, false);
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.Database.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            Polyline Poly2D = creeaza_poly2d(Data_table_centerline);
                            Polyline3d Poly3D = null;

                            if (project_type == "3d")
                            {

                                Poly3D = Functions.Build_3d_poly_for_scanning(Data_table_centerline);

                            }



                            string continut = "";

                            string sta_string = "";
                            string mp_string = "";
                            string offset_string = "";
                            string x_string = "";
                            string y_string = "";
                            string z_string = "";
                            string lat_string = "";
                            string long_string = "";

                            Point3d mleader_ins_pt;



                            double dist1 = new Point3d(picked_pt.X, picked_pt.Y, 0).DistanceTo(new Point3d(pt_on_poly.X, pt_on_poly.Y, 0));

                            offset_string = "Offset=" + Functions.Get_String_Rounded(dist1, 2);

                            if (dist1 > 0.009)
                            {
                                mleader_ins_pt = picked_pt;
                                Autodesk.AutoCAD.DatabaseServices.Line line1 = new Autodesk.AutoCAD.DatabaseServices.Line(new Point3d(picked_pt.X, picked_pt.Y, 0), new Point3d(pt_on_poly.X, pt_on_poly.Y, 0));
                                line1.Layer = layer_no_plot;
                                BTrecord.AppendEntity(line1);
                                Trans1.AddNewlyCreatedDBObject(line1, true);
                            }
                            else
                            {
                                mleader_ins_pt = pt_on_poly;
                            }

                            if (project_type == "2d")
                            {
                                if (Poly2D.Length < sta1 - start1)
                                {
                                    MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                    Freeze_operations = false;
                                    return;
                                }



                                sta_string = "STA2d=" + Functions.Get_chainage_from_double(sta1, "f", 2);
                                mp_string = "MP2d=" + Functions.Get_String_Rounded(mp1, 2);

                            }
                            else
                            {
                                if (Poly3D.Length < sta1 - start1)
                                {
                                    MessageBox.Show("the station you specified minus the starting station is larger than polyline length ");
                                    Freeze_operations = false;
                                    return;
                                }


                                sta_string = "STA3d=" + Functions.Get_chainage_from_double(sta1, "f", 2);
                                mp_string = "MP3d=" + Functions.Get_String_Rounded(mp1, 2);
                            }


                            x_string = "X=" + textBox_x.Text;
                            y_string = "Y=" + textBox_y.Text;
                            z_string = "Z=" + textBox_z.Text;
                            lat_string = "Lat=" + textBox_lat.Text;
                            long_string = "Long=" + textBox_long.Text;

                            if (checkBox_station.Checked == true)
                            {
                                continut = sta_string;
                            }

                            if (checkBox_mp.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = mp_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + mp_string;
                                }
                            }

                            if (checkBox_offset.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = offset_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + offset_string;
                                }
                            }

                            if (checkBox_xyz.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = x_string + " " + y_string + " " + z_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + x_string + " " + y_string + " " + z_string;
                                }
                            }



                            if (checkBox_ll.Checked == true)
                            {
                                if (continut == "")
                                {
                                    continut = lat_string + " " + long_string;
                                }
                                else
                                {
                                    continut = continut + "\r\n" + lat_string + " " + long_string;
                                }
                            }

                            if (checkBox_custom.Checked == true)
                            {
                                if (textBox_custom.Text != "")
                                {
                                    continut = textBox_custom.Text + "\r\n" + continut;
                                }
                            }


                            MLeader ml1 = Functions.creaza_mleader(mleader_ins_pt, continut, 10, 50, 200, 5, 10, 10);
                            ml1.Layer = layer_no_plot;


                            zoom_to_Point(mleader_ins_pt, 200);
                            if (project_type == "3d") delete_cl(Poly3D);
                            delete_cl(Poly2D);
                            Trans1.Commit();
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                Autodesk.Gis.Map.Platform.AcMapMap Acmap = Autodesk.Gis.Map.Platform.AcMapMap.GetCurrentMap();
                string Curent_system = Acmap.GetMapSRS();
                if (Curent_system != "")
                {
                    OSGeo.MapGuide.MgCoordinateSystemFactory Coord_factory1 = new OSGeo.MapGuide.MgCoordinateSystemFactory();
                    OSGeo.MapGuide.MgCoordinateSystem CoordSys1 = Coord_factory1.Create(Curent_system);
                    set_label_cs_to_green(CoordSys1.CsCode);
                }
                else
                {
                    set_label_cs_to_red();
                }


                delete_zoom_labels();
                Editor1.WriteMessage("\nCommand:");
                Freeze_operations = false;
            }
        }

        private void button_select_all_Click(object sender, EventArgs e)
        {
            if (dt_layer != null && dt_layer.Rows.Count > 0)
            {
                for (int i = 0; i < dt_layer.Rows.Count; ++i)
                {
                    dt_layer.Rows[i]["Select"] = true;
                }
            }
        }

        private void tabPage_shape_exp_Click(object sender, EventArgs e)
        {
            if (panel_dan.Visible == false)
            {
                panel_dan.Visible = true;
            }
            else
            {
                checkBox_user_select_export.Checked = false;
                panel_dan.Visible = false;

            }
        }
    }
    public sealed class Utility
    {
        private Utility()
        {
        }

        public static void ShowMsg(string msg)
        {
            AcadEditor.WriteMessage(msg);
        }

        public static Autodesk.AutoCAD.EditorInput.Editor AcadEditor
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            }
        }
    }


}

