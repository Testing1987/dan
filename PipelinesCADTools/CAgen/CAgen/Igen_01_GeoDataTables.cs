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


namespace Alignment_mdi
{
    public partial class Igen_geomanager : Form
    {

        string Tab = "\t";
        string vbcrlf = "\r\n";
        System.Data.DataTable Data_table_OD_attrib_existing;

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

        List<string> List_of_tables_on_layer;

        //int Oldw1 = 148;
        // int Oldl2 = 162;
        //int Oldw2 = 148;
        // int Oldll2 = 170;
        //int Oldl3 = 316;

        //overlap form


        public Igen_geomanager()
        {
            InitializeComponent();

            if (Environment.UserName.ToUpper() == "POP70694")
            {
                radioButton_BLOCKS.Visible = true;
                button_add_OD_table.Visible = true;
            }

            //Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute("._blockicon" + "\r\n", true, false, false);
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();

            ToolTip1.SetToolTip(this.button_add_OD_table, "All Objects on Selected Layer Will Be Assigned the Selected Object Data Table." + "\r\nData from the Selected Table Will Be Maintained. All Other Tables and Data Will Be Purged.");
            ToolTip1.SetToolTip(this.Button_Update_object_data, "Any Changes Made on Table Will Update Features in Drawing.");
            ToolTip1.SetToolTip(this.button_refresh_grid, "Refresh and Display All Features on the Selected Layer. Also Refreshes the Statistics Table.");
            ToolTip1.SetToolTip(this.button_zoom, "Zooms to Selected Row on Tables Corresponding Feature.");
            ToolTip1.SetToolTip(this.comboBox_od_existing_tables, "Specify Appropriate Data Table Based on Current Layer.");
            ToolTip1.SetToolTip(this.comboBox_layers_blocks_geomanager, "Specify Which Layer You Want to Load the Object Data Table.");
            ToolTip1.SetToolTip(this.button_refresh_layer_tables, "Load Layer Names & Object Data Tables from Current Drawing.");
            ToolTip1.SetToolTip(this.button_zoom_row_object_data, "Designate Geometry in Model Space You Want to Display on Table.");
            ToolTip1.SetToolTip(this.button_Filter, "Filter to Issues Identified in Statistics Table.");


            StopService("TabletInputService", 36000000);

            System.Diagnostics.Process[] proc = System.Diagnostics.Process.GetProcessesByName("tabtip");

            try
            {
                for (int i = 0; i < proc.Length; i = i + 1)
                {
                    proc[i].Kill();
                }
            }
            catch (System.Exception EX)
            {

            }

            panel_excel.Visible = false;



            // overlap form

            label_processing1.Visible = false;
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

                                //int Neww1 = GetMaxSize(List1, Oldw1);
                                //comboBox_layers_blocks_geomanager.Width = Neww1;

                                //comboBox_od_existing_tables.Left = Oldl2 + (Neww1 - Oldw1);

                                //label_correct_od_table.Left = Oldll2 + (Neww1 - Oldw1);

                                //int Neww2 = GetMaxSize(List2, Oldw2);
                                //comboBox_od_existing_tables.Width = Neww2;


                                //button_refresh_grid.Left = Oldl3 + (Neww1 - Oldw1) + (Neww2 - Oldw2);

                                button_refresh_grid.Top = 124;
                                comboBox_layers_blocks_geomanager.Top = 44;
                                label_current_layer_block.Top = 26;

                            }

                            if (radioButton_BLOCKS.Checked == true)
                            {

                                //int Neww1 = GetMaxSize(List1, Oldw1);
                                // comboBox_layers_blocks_geomanager.Width = Neww1;

                                //button_refresh_grid.Left = Oldl2 + (Neww1 - Oldw1);

                                button_refresh_grid.Top = 55;
                                comboBox_layers_blocks_geomanager.Top = 25;
                                label_current_layer_block.Top = 4;

                            }

                            label_drawing_name.Text = ThisDrawing.Database.Filename;
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
                if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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
                                                                            String Valoare1 = Record1[i].StrValue;
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
                                                    System.Data.DataTable Table1 = null; //Functions.Read_block_attributes_and_values(Block1);

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

                if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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
                                                    //Functions.Update_Attrib_block_values(Block1, Col_name, Col_value);

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

            if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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

            if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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

                if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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

                if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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

                                    Point3d minx = Ent1.GeometricExtents.MinPoint;
                                    Point3d maxx = Ent1.GeometricExtents.MaxPoint;

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
                                                    view.ZoomExtents(Ent1.GeometricExtents.MaxPoint, Ent1.GeometricExtents.MinPoint);
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

            if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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

                                dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

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

            if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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
                    if (Functions.Get_number_of_workbooks_from_Excel() > 1)
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

            if (ThisDrawing.Database.Filename != label_drawing_name.Text)
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
            }
            else
            {
                panel_excel.Visible = true;
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
                label_layer_rules.Visible = true;
                label_dt_issues.Visible = true;
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
                label_layer_rules.Visible = false;
                label_dt_issues.Visible = false;
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

                button_refresh_grid.Top = 124;
                comboBox_od_existing_tables.Visible = true;
                label_correct_od_table.Visible = true;
                comboBox_layers_blocks_geomanager.Top = 44;
                label_current_layer_block.Top = 26;


                // comboBox_layers_blocks_geomanager.Width = Oldw1;
                //comboBox_od_existing_tables.Left = Oldl2;
                //label_correct_od_table.Left = Oldll2;
                //button_refresh_grid.Left = Oldl3;
            }

            if (radioButton_BLOCKS.Checked == true)
            {

                button_refresh_grid.Top = 55;
                comboBox_od_existing_tables.Visible = false;
                label_correct_od_table.Visible = false;
                comboBox_layers_blocks_geomanager.Top = 25;
                label_current_layer_block.Top = 4;

                // comboBox_layers_blocks_geomanager.Width = Oldw1;

                //button_refresh_grid.Left = Oldl2;
            }

            this.Refresh();


        }


    }
}
