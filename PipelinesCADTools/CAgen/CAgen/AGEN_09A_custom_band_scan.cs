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
    public partial class AGEN_custom_band_scan : Form
    {





        public AGEN_custom_band_scan()
        {
            InitializeComponent();
        }


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_new_custom);
            lista_butoane.Add(button_open_excel_custom);
            lista_butoane.Add(button_scan_custom_data);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_custom_od_table);
            lista_butoane.Add(comboBox_custom_field1_od);
            lista_butoane.Add(comboBox_custom_field2_od);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_new_custom);
            lista_butoane.Add(button_open_excel_custom);
            lista_butoane.Add(button_scan_custom_data);
            lista_butoane.Add(comboBox_segment_name);
            lista_butoane.Add(comboBox_custom_od_table);
            lista_butoane.Add(comboBox_custom_field1_od);
            lista_butoane.Add(comboBox_custom_field2_od);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        private void TextBox_pt0_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_doubles_at_keypress(sender, e);



        }
        public string get_comboBox_band_excel_name()
        {
            return comboBox_band_excel_name.Text;
        }

        public void set_comboBox_band_excel_name(string txt)
        {
            if (txt != "")
            {
                if (comboBox_band_excel_name.Items.Contains(txt) == false)
                {
                    comboBox_band_excel_name.Items.Add(txt);
                }
                comboBox_band_excel_name.SelectedIndex = comboBox_band_excel_name.Items.IndexOf(txt);
            }
        }

        public string get_comboBox_custom_od_table()
        {
            return comboBox_custom_od_table.Text;
        }

        public void set_comboBox_custom_od_table(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_od_table.Items.Contains(txt) == false)
                {
                    comboBox_custom_od_table.Items.Add(txt);
                }
                comboBox_custom_od_table.SelectedIndex = comboBox_custom_od_table.Items.IndexOf(txt);
            }
        }

        public string get_comboBox_custom_field1_od()
        {
            return comboBox_custom_field1_od.Text;
        }

        public void set_comboBox_custom_field1_od(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_field1_od.Items.Contains(txt) == false)
                {
                    comboBox_custom_field1_od.Items.Add(txt);
                }
                comboBox_custom_field1_od.SelectedIndex = comboBox_custom_field1_od.Items.IndexOf(txt);
            }
        }

        public string get_comboBox_custom_field2_od()
        {
            return comboBox_custom_field2_od.Text;
        }

        public void set_comboBox_custom_field2_od(string txt)
        {
            if (txt != "")
            {
                if (comboBox_custom_field2_od.Items.Contains(txt) == false)
                {
                    comboBox_custom_field2_od.Items.Add(txt);
                }
                comboBox_custom_field2_od.SelectedIndex = comboBox_custom_field2_od.Items.IndexOf(txt);
            }
        }

        private void button_scan_custom_data_Click(object sender, EventArgs e)
        {
            if (comboBox_band_excel_name.Text == "")
            {
                MessageBox.Show("you did not specified the excel file name");
                return;
            }

            string custom_excel_name = comboBox_band_excel_name.Text + ".xlsx";

           
            if (Functions.Get_if_workbook_is_open_in_Excel(custom_excel_name) == true)
            {
                MessageBox.Show("Please close the " + custom_excel_name + " file");
                return;
            }

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }

            if (System.IO.Directory.Exists(ProjF) == true)
            {

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline data file does not exist");
                    _AGEN_mainform.dt_station_equation = null;
                    return;
                }

                if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                }



            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }




            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                set_enable_true();
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            string fisier_custom = ProjF + custom_excel_name;
            Functions.create_backup(fisier_custom);

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

         

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;

            ObjectId[] Empty_array = null;
            Editor1.SetImpliedSelection(Empty_array);
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            // Ag.WindowState = FormWindowState.Minimized;


            try
            {




                string od_field1 = comboBox_custom_field1_od.Text;
                string od_field2 = comboBox_custom_field2_od.Text;
                string custom_table_name = comboBox_custom_od_table.Text;






                int index_custom = -1;

                if (custom_table_name != "")
                {
                    if (_AGEN_mainform.Data_Table_custom_bands != null)
                    {
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                        {

                            for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                            {
                                if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                                {
                                    string bn = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);
                                    if (bn == custom_table_name)
                                    {
                                        index_custom = i;


                                        i = _AGEN_mainform.Data_Table_custom_bands.Rows.Count;
                                    }
                                }
                            }
                        }
                    }
                }


                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {

                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable1 = (Autodesk.AutoCAD.DatabaseServices.BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                        LayerTable Layer_table = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;

                        Autodesk.Gis.Map.ObjectData.Table Tabla0;

                        if (Tables1.IsTableDefined(custom_table_name) == true)
                        {
                            Tabla0 = Tables1[custom_table_name];

                            if (index_custom >= 0)
                            {
                                _AGEN_mainform.Data_Table_custom_bands.Rows[index_custom]["OD_table_name"] = custom_table_name;
                            }

                            bool field1_defined = false;
                            bool field2_defined = false;

                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla0.FieldDefinitions;
                            for (int i = 0; i < Field_defs1.Count; ++i)
                            {
                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                string Nume_field = Field_def1.Name;

                                if (od_field1 != "")
                                {
                                    if (Nume_field == od_field1)
                                    {

                                        if (index_custom >= 0) _AGEN_mainform.Data_Table_custom_bands.Rows[index_custom]["OD_field1"] = Nume_field;
                                    }
                                }

                                if (od_field2 != "")
                                {
                                    if (Nume_field == od_field2)
                                    {
                                        if (index_custom >= 0) _AGEN_mainform.Data_Table_custom_bands.Rows[index_custom]["OD_field2"] = Nume_field;
                                    }
                                }
                            }

                            if (index_custom >= 0)
                            {

                                if (field1_defined == false) _AGEN_mainform.Data_Table_regular_bands.Rows[index_custom]["OD_field1"] = DBNull.Value;
                                if (field2_defined == false) _AGEN_mainform.Data_Table_regular_bands.Rows[index_custom]["OD_field2"] = DBNull.Value;
                            }
                        }
                        else
                        {
                            set_enable_true();
                            MessageBox.Show("please specify an object data table!");
                            _AGEN_mainform.tpage_processing.Hide();
                            return;
                        }


                        string Layer_custom = "";

                        foreach (ObjectId ObjID in BTrecord)
                        {
                            Entity Ent_intersection = Trans1.GetObject(ObjID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as Entity;
                            if (Ent_intersection != null && Ent_intersection.IsErased == false)
                            {
                                LayerTableRecord Layer_rec = Trans1.GetObject(Layer_table[Ent_intersection.Layer], OpenMode.ForRead) as LayerTableRecord;
                                if (Ent_intersection is Polyline && Layer_rec.IsOff == false && Layer_rec.IsFrozen == false)
                                {
                                    Polyline Poly_int = Ent_intersection as Polyline;
                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla0.GetObjectTableRecords(Convert.ToUInt32(0), Ent_intersection.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, false))
                                    {
                                        if (Records1 != null)
                                        {
                                            Layer_custom = Ent_intersection.Layer;
                                            goto jump;
                                        }
                                    }
                                }
                            }
                        }

                    jump:

                        if (Layer_custom == "")
                        {

                            MessageBox.Show("The object data table you specified is not attached to any parcel polyline in the drawing\r\nmaybe you shoud map import shp current....");
                            Ag.WindowState = FormWindowState.Normal;
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            return;
                        }

                        System.Data.DataTable dt1 = Functions.Creaza_custom_datatable_structure(od_field1, od_field2);

                        System.Data.DataTable dt_custom = Functions.Scan_parcels(_AGEN_mainform.dt_centerline, _AGEN_mainform.dt_station_equation, dt1, Layer_custom, _AGEN_mainform.Project_type, custom_table_name, od_field1, od_field2, od_field1, od_field2);


                        if (dt_custom != null)
                        {
                            if (dt_custom.Rows.Count > 0)
                            {
                                Populate_custom_file(fisier_custom, dt_custom,_AGEN_mainform.config_path);
                            }
                        }






                        Trans1.Commit();
                    }
                }


                ThisDrawing.Editor.WriteMessage("\nCommand:");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            _AGEN_mainform.tpage_processing.Hide();


            set_enable_true();
            Ag.WindowState = FormWindowState.Normal;

        }

        public void Populate_custom_file(string file_custom, System.Data.DataTable dt_custom, string cfg_file)
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

                if (System.IO.File.Exists(file_custom) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }
                else
                {
                    Workbook1 = Excel1.Workbooks.Open(file_custom);
                }
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                Microsoft.Office.Interop.Excel.Workbook Workbook2 = Excel1.Workbooks.Open(cfg_file);
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                string segment1 = _AGEN_mainform.current_segment;
                if (segment1 == "not defined") segment1 = "";

                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh2 in Workbook2.Worksheets)
                {
                    if (wsh2.Name == comboBox_band_excel_name.Text + "_cfg_" + segment1)
                    {
                        W2 = wsh2;
                    }
                }

                if (W2 == null)
                {
                    W2 = Workbook2.Worksheets.Add(System.Reflection.Missing.Value, Workbook2.Worksheets[Workbook2.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    if ((comboBox_band_excel_name.Text + "_cfg_" + segment1).Length > 31)
                    {
                        MessageBox.Show(comboBox_band_excel_name.Text + "_cfg_" + segment1 + "is bigger than 31 charcaters\r\nor you rename the custom band to have less characters\r\nor/and rename the segment");
                        return;
                    }

                    W2.Name = comboBox_band_excel_name.Text + "_cfg_" + segment1;
                }

                try
                {

                    Functions.Transfer_to_worksheet_Data_table(W1, dt_custom, _AGEN_mainform.Start_row_custom, "General");
                    Functions.Create_header_custom_file(W1, comboBox_band_excel_name.Text, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    if (System.IO.File.Exists(file_custom) == false)
                    {
                        Workbook1.SaveAs(file_custom);
                    }
                    else
                    {
                        Workbook1.Save();
                    }
                    Workbook1.Close();


                    int NrR = 10;
                    int NrC = 2;



                    Object[,] values = new object[NrR, NrC];
                    values[0, 0] = "Band Excel File Name";
                    values[1, 0] = "OD Table";
                    values[2, 0] = "OD Field1";
                    values[3, 0] = "OD Field2";

                    values[4, 0] = "Custom Band Block";
                    values[5, 0] = "Block Tag Sta1";
                    values[6, 0] = "Block Tag Sta2";
                    values[7, 0] = "Block Tag Length";
                    values[8, 0] = "Block Tag Attribute 1";
                    values[9, 0] = "Block Tag Attribute 2";

                    values[0, 1] = get_comboBox_band_excel_name();
                    values[1, 1] = get_comboBox_custom_od_table();
                    values[2, 1] = get_comboBox_custom_field1_od();
                    values[3, 1] = get_comboBox_custom_field2_od();



                    Microsoft.Office.Interop.Excel.Range range2 = W2.Range["A1:B10"];
                    range2.Cells.NumberFormat = "General";
                    range2.Value2 = values;
                    Functions.Color_border_range_inside(range2, 0);

                    Workbook2.Save();

                    _AGEN_mainform.dt_settings_custom = null;

                    foreach (Microsoft.Office.Interop.Excel.Worksheet W3 in Workbook2.Worksheets)
                    {
                        try
                        {
                            #region build Custom_datatable_config
                            if (W3.Name.Contains("_cfg_" + segment1) == true)
                            {
                                _AGEN_mainform.tpage_setup.build_dt_custom_settings(W3);
                            }
                            #endregion
                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                    }

                    Workbook2.Close();




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
                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }




        private void comboBox_custom_od_table_SelectedIndexChanged(object sender, EventArgs e)
        {
            Functions.add_OD_fieds_to_combobox(comboBox_custom_od_table, comboBox_custom_field1_od);
            Functions.add_OD_fieds_to_combobox(comboBox_custom_od_table, comboBox_custom_field2_od);
        }



        private void button_show_draw_custom_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
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

            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_cust_draw.Show();

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;


        }

        private void button_refresh_bands_Click(object sender, EventArgs e)
        {
            if (_AGEN_mainform.Data_Table_custom_bands != null)
            {
                if (_AGEN_mainform.Data_Table_custom_bands.Rows.Count > 0)
                {
                    comboBox_band_excel_name.Items.Clear();
                    comboBox_band_excel_name.Items.Add("");

                    for (int i = 0; i < _AGEN_mainform.Data_Table_custom_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_custom_bands.Rows[i]["band_name"]);

                            comboBox_band_excel_name.Items.Add(bn);


                        }
                    }

                    comboBox_band_excel_name.SelectedIndex = 0;
                }
            }
        }

        private void comboBox_excel_name_SelectedIndexChanged(object sender, EventArgs e)
        {

            string band_name = comboBox_band_excel_name.Text;

            if (band_name != "")
            {
                if (_AGEN_mainform.dt_settings_custom != null)
                {
                    if (_AGEN_mainform.dt_settings_custom.Rows.Count > 0)
                    {

                        for (int i = 0; i < _AGEN_mainform.dt_settings_custom.Rows.Count; ++i)
                        {
                            if (_AGEN_mainform.dt_settings_custom.Rows[i][0] != DBNull.Value)
                            {
                                string bn = Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][0]);
                                if (bn == band_name)
                                {
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][1] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_od_table, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][1]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][2] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_field1_od, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][2]));
                                    }
                                    if (_AGEN_mainform.dt_settings_custom.Rows[i][3] != DBNull.Value)
                                    {
                                        add_to_combobox(comboBox_custom_field2_od, Convert.ToString(_AGEN_mainform.dt_settings_custom.Rows[i][3]));
                                    }

                                    i = _AGEN_mainform.dt_settings_custom.Rows.Count;
                                }


                            }
                        }
                    }
                }
            }

        }

        private void add_to_combobox(System.Windows.Forms.ComboBox combo1, string string1)
        {
            if (combo1.Items.Contains(string1) == false)
            {
                combo1.Items.Add(string1);
            }
            combo1.SelectedIndex = combo1.Items.IndexOf(string1);
        }

        public void clear_combobox_custom()
        {
            comboBox_band_excel_name.Items.Clear();
            comboBox_custom_od_table.Items.Clear();
            comboBox_custom_field1_od.Items.Clear();
            comboBox_custom_field2_od.Items.Clear();

        }

        private void button_open_excel_custom_Click(object sender, EventArgs e)
        {
           
           
               
                try
                {
                    string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                    if (System.IO.Directory.Exists(ProjF) == true)
                    {

                        if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                        {
                            ProjF = ProjF + "\\";
                        }
                        string custom_excel_name = comboBox_band_excel_name.Text + ".xlsx";
                        string fisier_custom = ProjF + custom_excel_name;

                        if (System.IO.File.Exists(fisier_custom) == false)
                        {
                            
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
                        Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(fisier_custom);
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
          

           
        }

        private void button_Load_od_Click(object sender, EventArgs e)
        {


            set_enable_false();
                try
                {
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                    using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                    {
                        using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                        {
                            comboBox_custom_od_table.Items.Clear();
                            comboBox_custom_field1_od.Items.Clear();
                            comboBox_custom_field2_od.Items.Clear();


                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            System.Collections.Specialized.StringCollection Nume_tables = new System.Collections.Specialized.StringCollection();
                            Nume_tables = Tables1.GetTableNames();

                            for (int i = 0; i < Nume_tables.Count; i = i + 1)
                            {
                                String Tabla1 = Nume_tables[i];
                                if (comboBox_custom_od_table.Items.Contains(Tabla1) == false)
                                {
                                    comboBox_custom_od_table.Items.Add(Tabla1);
                                }
                            }


                            this.Refresh();
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
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

        public void set_combobox_segment_name()
        {
            comboBox_segment_name.SelectedIndex = comboBox_segment_name.Items.IndexOf(_AGEN_mainform.current_segment);
        }

        private void ComboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            _AGEN_mainform.current_segment = comboBox_segment_name.Text;
            _AGEN_mainform.tpage_setup.set_combobox_segment_name();


        }

        private void button_new_custom_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W1 = null;
            try
            {
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    Excel1.Visible = true;
                }
                catch (System.Exception)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    Excel1.Visible = true;
                }

                Workbook1 = Excel1.Workbooks.Add();
                W1 = Workbook1.Worksheets[1];

                System.Data.DataTable dt1 = Functions.Creaza_custom_datatable_structure("Field1", "Field2");
                Functions.Create_header_custom_file(W1, comboBox_band_excel_name.Text, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), _AGEN_mainform.current_segment);
                for (int i = 0; i < dt1.Columns.Count; ++i)
                {
                    string letter1 = Functions.get_excel_column_letter(i + 1);
                    W1.Range[letter1 + "8"].Value2 = dt1.Columns[i].ColumnName;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }

}


