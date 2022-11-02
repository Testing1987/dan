using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class SGEN_Settings : Form
    {
        public static SGEN_Settings tpage_settings = null;

        public static string current_segment = "";

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();


            lista_butoane.Add(radioButton_Load_config);
            lista_butoane.Add(radioButton_new_config);
            lista_butoane.Add(button_browser_dwt);
            lista_butoane.Add(button_browse_select_output_folder);
            lista_butoane.Add(comboBox_vw_scale);
            lista_butoane.Add(comboBox_dwgunits);
            lista_butoane.Add(button_pick_plan_view);


            lista_butoane.Add(dataGridView_bands);
            lista_butoane.Add(button_align_config_saveall);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(radioButton_Load_config);
            lista_butoane.Add(radioButton_new_config);
            lista_butoane.Add(button_browser_dwt);
            lista_butoane.Add(button_browse_select_output_folder);
            lista_butoane.Add(comboBox_vw_scale);
            lista_butoane.Add(comboBox_dwgunits);
            lista_butoane.Add(button_pick_plan_view);


            lista_butoane.Add(dataGridView_bands);
            lista_butoane.Add(button_align_config_saveall);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public SGEN_Settings()
        {
            InitializeComponent();
            _SGEN_mainform.dt_segments = new System.Data.DataTable();
            _SGEN_mainform.dt_segments.Columns.Add("Template", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Output folder", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Prefix File Name", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Suffix File Name", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Start numbering", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Increment", typeof(string));
            _SGEN_mainform.dt_segments.Columns.Add("Segment Name", typeof(string));



        }

        private void button_load_config_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog fbd = new OpenFileDialog())
            {
                fbd.Multiselect = false;
                fbd.Filter = "Excel files (*.xlsx)|*.xlsx";

                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {



                    _SGEN_mainform.config_file = fbd.FileName;
                    _SGEN_mainform.project_main_folder = System.IO.Path.GetDirectoryName(_SGEN_mainform.config_file);

                    radioButton_Load_config.Checked = true;

                    _SGEN_mainform.dt_segments = new System.Data.DataTable();
                    _SGEN_mainform.dt_segments.Columns.Add("Template", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Output folder", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Prefix File Name", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Suffix File Name", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Start numbering", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Increment", typeof(string));
                    _SGEN_mainform.dt_segments.Columns.Add("Segment Name", typeof(string));


                    #region Load_config_method
                    {
                        set_enable_false();
                        Load_existing_config_file(_SGEN_mainform.config_file);
                        set_enable_true();


                        if (_SGEN_mainform.dt_segments.Rows.Count > 0)
                        {
                            if (_SGEN_mainform.dt_segments.Rows[0]["Segment Name"] != DBNull.Value)
                            {
                                string segment1 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[0]["Segment Name"]);

                                populate_controls_on_page(segment1);

                            }

                            Fill_combobox_segments();
                        }

                    }
                    #endregion






                }
            }
        }


        private void Load_existing_config_file(string File1)
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);

                try
                {
                    int no_worksheets = Workbook1.Worksheets.Count;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        try
                        {
                            #region main_worksheet

                            if (W1.Name == "main_cfg")
                            {

                                string b1 = Convert.ToString(W1.Range["B1"].Value2);
                                string b2 = Convert.ToString(W1.Range["B2"].Value);
                                string b3 = Convert.ToString(W1.Range["B3"].Value);
                                string b4 = Convert.ToString(W1.Range["B4"].Value);
                                string b5 = Convert.ToString(W1.Range["B5"].Value);
                                string b6 = Convert.ToString(W1.Range["B6"].Value);
                                string b7 = Convert.ToString(W1.Range["B7"].Value);
                                string b8 = Convert.ToString(W1.Range["B8"].Value);
                                string b9 = Convert.ToString(W1.Range["B9"].Value);
                                string b10 = Convert.ToString(W1.Range["B10"].Value);
                                string b11 = Convert.ToString(W1.Range["B11"].Value);
                                string b12 = Convert.ToString(W1.Range["B12"].Value);
                                string b13 = Convert.ToString(W1.Range["B13"].Value);
                                string b14 = Convert.ToString(W1.Range["B14"].Value);
                                string b15 = Convert.ToString(W1.Range["B15"].Value);
                                string b16 = Convert.ToString(W1.Range["B16"].Value);
                                string b17 = Convert.ToString(W1.Range["B17"].Value);
                                string b18 = Convert.ToString(W1.Range["B18"].Value);
                                string b19 = Convert.ToString(W1.Range["B19"].Value);
                                string b20 = Convert.ToString(W1.Range["B20"].Value);
                                string b21 = Convert.ToString(W1.Range["B21"].Value);
                                string b22 = Convert.ToString(W1.Range["B22"].Value);
                                string b23 = Convert.ToString(W1.Range["B23"].Value);
                                string b24 = Convert.ToString(W1.Range["B24"].Value);
                                string b25 = Convert.ToString(W1.Range["B25"].Value);
                                string b26 = Convert.ToString(W1.Range["B26"].Value);
                                string b27 = Convert.ToString(W1.Range["B27"].Value);
                                string b28 = Convert.ToString(W1.Range["B28"].Value);
                                string b29 = Convert.ToString(W1.Range["B29"].Value);
                                string b30 = Convert.ToString(W1.Range["B30"].Value);
                                string b31 = Convert.ToString(W1.Range["B31"].Value);
                                string b32 = Convert.ToString(W1.Range["B32"].Value);
                                string b33 = Convert.ToString(W1.Range["B33"].Value);
                                string b34 = Convert.ToString(W1.Range["B34"].Value);
                                string b35 = Convert.ToString(W1.Range["B35"].Value);
                                string b36 = Convert.ToString(W1.Range["B36"].Value);
                                string b37 = Convert.ToString(W1.Range["B37"].Value);
                                string b38 = Convert.ToString(W1.Range["B38"].Value);
                                string b39 = Convert.ToString(W1.Range["B39"].Value);
                                string b40 = Convert.ToString(W1.Range["B40"].Value);
                                string b41 = Convert.ToString(W1.Range["B41"].Value);
                                string b42 = Convert.ToString(W1.Range["B42"].Value);
                                string b43 = Convert.ToString(W1.Range["B43"].Value);
                                string b44 = Convert.ToString(W1.Range["B44"].Value);
                                string b45 = Convert.ToString(W1.Range["B45"].Value);
                                string b46 = Convert.ToString(W1.Range["B46"].Value);
                                string b47 = Convert.ToString(W1.Range["B47"].Value);

                                string c4 = Convert.ToString(W1.Range["C4"].Value);
                                string c5 = Convert.ToString(W1.Range["C5"].Value);
                                string c6 = Convert.ToString(W1.Range["C6"].Value);
                                string c7 = Convert.ToString(W1.Range["C7"].Value);
                                string c8 = Convert.ToString(W1.Range["C8"].Value);
                                string c9 = Convert.ToString(W1.Range["C9"].Value);
                                string c10 = Convert.ToString(W1.Range["C10"].Value);

                                if (b1 != null) textBox_client_name.Text = b1.ToString();
                                if (b2 != null) textBox_project_name.Text = b2.ToString();

                                if (b3 != null)
                                {
                                    if (Functions.IsNumeric(b3.ToString()) == true)
                                    {
                                        _SGEN_mainform.no_of_segments = Convert.ToInt32(b3.ToString());
                                    }
                                    else
                                    {
                                        _SGEN_mainform.no_of_segments = 0;
                                    }

                                }
                                else
                                {
                                    _SGEN_mainform.no_of_segments = 0;
                                }

                                if (_SGEN_mainform.no_of_segments > 0)
                                {
                                    if (c10 != null)
                                    {
                                        label_sheet_naming.Text = "Sheet Setup " + c10.ToString();
                                    }
                                    else
                                    {
                                        label_sheet_naming.Text = "Sheet Setup";
                                        MessageBox.Show("you can't have the cell B3 specifying " + Convert.ToString(_SGEN_mainform.no_of_segments) + " segments and on C10 not to have specified a segment name\r\nFix the issue!");
                                        return;
                                    }
                                }
                                else
                                {
                                    label_sheet_naming.Text = "Sheet Setup";
                                }

                                if (_SGEN_mainform.no_of_segments > 0)
                                {
                                    string start_cell = "C4";

                                    string end_cell = Functions.get_excel_column_letter(_SGEN_mainform.no_of_segments + 2) + "10";


                                    for (int i = 0; i < _SGEN_mainform.no_of_segments; ++i)
                                    {
                                        _SGEN_mainform.dt_segments.Rows.Add();
                                    }


                                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[start_cell + ":" + end_cell];

                                    object[,] values = new object[_SGEN_mainform.dt_segments.Columns.Count - 1, _SGEN_mainform.no_of_segments - 1];

                                    values = range1.Value2;


                                    for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                                    {
                                        for (int j = 0; j < _SGEN_mainform.dt_segments.Columns.Count; ++j)
                                        {
                                            object Valoare = values[j + 1, i + 1];
                                            if (Valoare == null) Valoare = DBNull.Value;
                                            _SGEN_mainform.dt_segments.Rows[i][j] = Valoare;
                                        }

                                    }
                                }

                                if (b4 != null)
                                {
                                    string Template = b4.ToString();
                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c4 != null)
                                        {
                                            Template = c4.ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("you can't have the cell B3 specifying " + Convert.ToString(_SGEN_mainform.no_of_segments) + " segments and on C4 not to have specified a template\r\nFix the issue!");
                                        }
                                    }

                                    if (System.IO.File.Exists(Template) == true)
                                    {
                                        set_textBox_template_name(Template);
                                    }
                                }

                                if (b5 != null)
                                {
                                    string Output = b5.ToString();
                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c5 != null)
                                        {
                                            Output = c5.ToString();
                                        }
                                        else
                                        {
                                            MessageBox.Show("you can't have the cell B3 specifying " + Convert.ToString(_SGEN_mainform.no_of_segments) + " segments and on C5 not to have specified an output folder\r\nFix the issue!");
                                        }
                                    }
                                    if (System.IO.Directory.Exists(Output) == true)
                                    {
                                        Set_output_folder_text_box(Output);
                                    }
                                }

                                if (b6 != null)
                                {
                                    string Pref1 = b6.ToString();
                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c6 != null)
                                        {
                                            Pref1 = c6.ToString();
                                        }
                                    }
                                    if (Pref1 != null)
                                    {
                                        Set_prefix_text_box(Pref1);
                                    }
                                }

                                if (b7 != null)
                                {
                                    string Suffix1 = b7.ToString();

                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c7 != null)
                                        {
                                            Suffix1 = c7.ToString();
                                        }
                                    }


                                    Set_suffix_text_box(Suffix1);
                                }

                                if (b8 != null)
                                {
                                    string Startno = b8.ToString();


                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c8 != null)
                                        {
                                            Startno = c8.ToString();
                                        }
                                    }


                                    if (Functions.IsNumeric(Startno) == true)
                                    {
                                        Set_start_no_text_box(Startno);
                                    }
                                }

                                if (b9 != null)
                                {
                                    string Increment = b9.ToString();

                                    if (_SGEN_mainform.no_of_segments > 0)
                                    {
                                        if (c9 != null)
                                        {
                                            Increment = c9.ToString();
                                        }
                                    }

                                    if (Functions.IsNumeric(Increment) == true)
                                    {
                                        Set_increment_text_box(Increment);
                                    }
                                }




                                if (b15 != null && b15 == "meters")
                                {
                                    _SGEN_mainform.units_of_measurement = "m";
                                    Set_combobox_units_to_m();
                                }
                                else
                                {
                                    _SGEN_mainform.units_of_measurement = "f";
                                    Set_combobox_units_to_ft();
                                }

                                if (b16 != null)
                                {
                                    _SGEN_mainform.project_main_folder = b16.ToString();

                                    if (System.IO.Directory.Exists(_SGEN_mainform.project_main_folder) == true)
                                    {


                                        if (_SGEN_mainform.project_main_folder.Substring(_SGEN_mainform.project_main_folder.Length - 1, 1) != "\\")
                                        {
                                            _SGEN_mainform.project_main_folder = _SGEN_mainform.project_main_folder + "\\";
                                        }

                                        string projF = _SGEN_mainform.project_main_folder;
                                        if (_SGEN_mainform.no_of_segments > 0 && c10 != null)
                                        {
                                            projF = projF + c10.ToString();
                                        }

                                        if (System.IO.Directory.Exists(projF) == true)
                                        {
                                            Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;
                                            Microsoft.Office.Interop.Excel.Worksheet W2 = null;
                                            try
                                            {
                                                string fisier_si = projF + _SGEN_mainform.sheet_index_excel_name;
                                                if (System.IO.File.Exists(fisier_si) == true)
                                                {
                                                    bool is_sheet_index_open = false;
                                                    foreach (Microsoft.Office.Interop.Excel.Workbook wk1 in Excel1.Workbooks)
                                                    {
                                                        if (wk1.FullName.ToLower() == fisier_si.ToLower())
                                                        {
                                                            Workbook2 = wk1;
                                                            is_sheet_index_open = true;
                                                        }
                                                    }
                                                    if (Workbook2 == null) Workbook2 = Excel1.Workbooks.Open(fisier_si);
                                                    
                                                    W2 = Workbook2.Worksheets[1];
                                                    _SGEN_mainform.dt_sheet_index = _SGEN_mainform.tpage_sheetindex.Build_Data_table_sheet_index_from_excel(W2, _SGEN_mainform.Start_row_Sheet_index + 1);
                                                    _SGEN_mainform.tpage_sheetindex.set_dataGridView_sheet_index();
                                                    if (is_sheet_index_open == false) Workbook2.Close();
                                                }
                                            }
                                            catch (System.Exception ex)
                                            {
                                                System.Windows.Forms.MessageBox.Show(ex.Message);

                                            }
                                            finally
                                            {
                                                if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                                                if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook2);
                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("the Project database folder location is not specified\r\n" + projF + "\r\n operation aborted");
                                            _SGEN_mainform.project_main_folder = "";
                                            return;
                                        }


                                    }

                                }
                            }
                            #endregion

                            #region Regular_band_data
                            else if (W1.Name == "Regular_band_data")
                            {
                                _SGEN_mainform.Data_Table_regular_bands = build_regular_band_data_table_from_excel(W1, 2);
                                string main_vp_name = get_comboBox_viewport_target_areas();


                                if (_SGEN_mainform.Data_Table_regular_bands != null)
                                {
                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                        {
                                            if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                            {
                                                string bn = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);

                                                if (bn == main_vp_name)
                                                {
                                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] != DBNull.Value)
                                                    {
                                                        string str_scale = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"]);
                                                        if (Functions.IsNumeric(str_scale) == true)
                                                        {
                                                            _SGEN_mainform.Vw_scale = Convert.ToDouble(str_scale);



                                                            for (int k = 0; k < Get_combobox_viewport_scale_count(); ++k)
                                                            {
                                                                string Scale1 = Get_combobox_viewport_scale(k);

                                                                if (Scale1.Contains(":") == true)
                                                                {
                                                                    Scale1 = Scale1.Substring(2, Scale1.Length - 2);
                                                                    if (Functions.IsNumeric(Scale1) == true)
                                                                    {
                                                                        if (Math.Round(_SGEN_mainform.Vw_scale, 1) == Math.Round(1000 / Convert.ToDouble(Scale1), 1))
                                                                        {
                                                                            Set_combobox_viewport_scale(k);
                                                                            k = Get_combobox_viewport_scale_count();
                                                                        }
                                                                    }
                                                                }
                                                                string feet = "\u0022";

                                                                if (Scale1.Contains(feet + "=") == true && Scale1.Contains("'") == true)
                                                                {
                                                                    Scale1 = Scale1.Substring(3, Scale1.Length - 4);
                                                                    if (Functions.IsNumeric(Scale1) == true)
                                                                    {
                                                                        if (Math.Round(_SGEN_mainform.Vw_scale, 4) == Math.Round(1 / Convert.ToDouble(Scale1), 4))
                                                                        {
                                                                            Set_combobox_viewport_scale(k);
                                                                            k = Get_combobox_viewport_scale_count();
                                                                        }
                                                                    }
                                                                }
                                                            }

                                                            if (_SGEN_mainform.Vw_scale == 1)
                                                            {
                                                                Set_combobox_viewport_scale(0);
                                                            }

                                                        }
                                                    }

                                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _SGEN_mainform.Vw_ps_x = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _SGEN_mainform.Vw_ps_y = Convert.ToDouble(str_val);
                                                        }
                                                    }


                                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _SGEN_mainform.Vw_width = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                    if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] != DBNull.Value)
                                                    {
                                                        string str_val = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                                        if (Functions.IsNumeric(str_val) == true)
                                                        {
                                                            _SGEN_mainform.Vw_height = Convert.ToDouble(str_val);
                                                        }
                                                    }

                                                }








                                            }
                                        }
                                    }
                                }

                            }
                            #endregion

                        }
                        catch (System.Exception ex)
                        {
                            System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);

                }



                creeaza_display_data_table(Creaza_lista_regular_vp_picked());


                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        public void set_textBox_template_name(string dwt_name)
        {
            textBox_template_name.Text = dwt_name;
        }

        public void Set_output_folder_text_box(string output_f)
        {
            textBox_output_folder.Text = output_f;
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

        public static System.Data.DataTable creeaza_regular_band_data_table_structure()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("band_name", typeof(string));
            dt.Columns.Add("Custom_scale", typeof(double));
            dt.Columns.Add("OD_table_name", typeof(string));
            dt.Columns.Add("OD_field1", typeof(string));
            dt.Columns.Add("OD_field2", typeof(string));
            dt.Columns.Add("block_name", typeof(string));
            dt.Columns.Add("block_sta_atr1", typeof(string));
            dt.Columns.Add("block_sta_atr2", typeof(string));
            dt.Columns.Add("block_len_atr", typeof(string));
            dt.Columns.Add("block_field1", typeof(string));
            dt.Columns.Add("block_field2", typeof(string));
            dt.Columns.Add("band_separation", typeof(double));
            dt.Columns.Add("viewport_width", typeof(double));
            dt.Columns.Add("viewport_height", typeof(double));
            dt.Columns.Add("viewport_ps_x", typeof(double));
            dt.Columns.Add("viewport_ps_y", typeof(double));
            dt.Columns.Add("viewport_ms_x", typeof(double));
            dt.Columns.Add("viewport_ms_y", typeof(double));
            dt.Columns.Add("viewport_twist", typeof(double));
            return dt;
        }

        public System.Data.DataTable build_regular_band_data_table_from_excel(Microsoft.Office.Interop.Excel.Worksheet W1, int Start_row)
        {

            System.Data.DataTable dt_regular = creeaza_regular_band_data_table_structure();
            int NrR = 0;
            int NrC = dt_regular.Columns.Count;


            bool is_data = false;

            for (int i = Start_row; i < 30000; ++i)
            {
                if (i == Start_row)
                {
                    if (W1.Range["A" + i.ToString()].Value2 == null)
                    {
                        return dt_regular;
                    }
                }

                if (W1.Range["A" + i.ToString()].Value2 == null)
                {
                    NrR = i - Start_row;
                    i = 31000;
                }
                else
                {
                    dt_regular.Rows.Add();
                    is_data = true;
                }
            }

            if (is_data == true)
            {
                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[Start_row, 1], W1.Cells[NrR + Start_row - 1, NrC]];

                object[,] values = new object[NrR - 1, NrC - 1];

                values = range1.Value2;

                for (int i = 0; i < dt_regular.Rows.Count; ++i)
                {
                    for (int j = 0; j < dt_regular.Columns.Count; ++j)
                    {
                        object Valoare = values[i + 1, j + 1];
                        if (Valoare == null) Valoare = DBNull.Value;
                        dt_regular.Rows[i][j] = Valoare;
                    }
                }
            }

            if (dt_regular.Rows.Count > 0)
            {
                for (int i = 0; i < dt_regular.Rows.Count; ++i)
                {
                    if (Convert.ToString(dt_regular.Rows[i]["band_name"]) == _SGEN_mainform.nume_main_vp)
                    {
                        if (dt_regular.Rows[i]["band_separation"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_regular.Rows[i]["band_separation"])) == true)
                        {
                            _SGEN_mainform.Band_Separation = Convert.ToDouble(dt_regular.Rows[i]["band_separation"]);
                        }
                        if (dt_regular.Rows[i]["viewport_width"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_regular.Rows[i]["viewport_width"])) == true)
                        {
                            _SGEN_mainform.Vw_width = Convert.ToDouble(dt_regular.Rows[i]["viewport_width"]);
                        }
                        if (dt_regular.Rows[i]["viewport_height"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_regular.Rows[i]["viewport_height"])) == true)
                        {
                            _SGEN_mainform.Vw_height = Convert.ToDouble(dt_regular.Rows[i]["viewport_height"]);
                        }
                        if (dt_regular.Rows[i]["viewport_ps_x"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_regular.Rows[i]["viewport_ps_x"])) == true)
                        {
                            _SGEN_mainform.Vw_ps_x = Convert.ToDouble(dt_regular.Rows[i]["viewport_ps_x"]);
                        }
                        if (dt_regular.Rows[i]["viewport_ps_y"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_regular.Rows[i]["viewport_ps_y"])) == true)
                        {
                            _SGEN_mainform.Vw_ps_y = Convert.ToDouble(dt_regular.Rows[i]["viewport_ps_y"]);
                        }
                    }

                }
            }



            return dt_regular;

        }

        public string get_comboBox_viewport_target_areas()
        {
            return "Plan View";
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

        public static List<string> Creaza_lista_regular_vp_picked()
        {
            List<string> lista1 = new List<string>();
            if (_SGEN_mainform.Data_Table_regular_bands != null)
            {
                if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {

                        if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] != DBNull.Value && _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] != DBNull.Value)
                        {
                            string y_string = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"]);
                            string bandh_string = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                            if (Functions.IsNumeric(y_string) == true && Functions.IsNumeric(bandh_string) == true)
                            {
                                lista1.Add("YES");
                            }
                            else
                            {
                                lista1.Add("NO");
                            }

                        }
                        else
                        {
                            lista1.Add("NO");
                        }
                    }
                }
            }

            return lista1;
        }
        public void creeaza_display_data_table(List<string> Lista1)
        {
            _SGEN_mainform.Data_Table_display_bands = new System.Data.DataTable();
            _SGEN_mainform.Data_Table_display_bands.Columns.Add("Band Name", typeof(string));
            _SGEN_mainform.Data_Table_display_bands.Columns.Add("Location Selected", typeof(string));

            if (_SGEN_mainform.Data_Table_regular_bands != null && _SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
            {
                if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count == Lista1.Count)
                {
                    for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);

                            _SGEN_mainform.Data_Table_display_bands.Rows.Add();
                            _SGEN_mainform.Data_Table_display_bands.Rows[_SGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Band Name"] = bn;
                            _SGEN_mainform.Data_Table_display_bands.Rows[_SGEN_mainform.Data_Table_display_bands.Rows.Count - 1]["Location Selected"] = Lista1[i];
                        }
                    }
                }
            }
            dataGridView_bands.DataSource = _SGEN_mainform.Data_Table_display_bands;
            dataGridView_bands.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dataGridView_bands.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_bands.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_bands.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_bands.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_bands.EnableHeadersVisualStyles = false;
        }

        private void comboBox_dwgunits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[0].ToString())
            {
                _SGEN_mainform.units_of_measurement = "f";
            }
            if (comboBox_dwgunits.Text == comboBox_dwgunits.Items[1].ToString())
            {
                _SGEN_mainform.units_of_measurement = "m";
            }
            Set_content_of_combobox_viewport_scale();
        }

        public void Set_content_of_combobox_viewport_scale()
        {
            comboBox_vw_scale.Items.Clear();
            if (_SGEN_mainform.units_of_measurement == "f")
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

        private void button_browse_select_output_folder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_output_folder.Text = fbd.SelectedPath.ToString();
                    _SGEN_mainform.output_folder = textBox_output_folder.Text;
                }

            }
        }

        private void button_pick_plan_view_Click(object sender, EventArgs e)
        {
            set_enable_false();
            try
            {

                string band_name = "";

                if (System.IO.File.Exists(_SGEN_mainform.config_file) == true)
                {






                    string strTemplatePath = get_template_name_from_text_box();
                    DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = null;

                    if (System.IO.File.Exists(strTemplatePath) == false)
                    {
                        MessageBox.Show("template file not found");
                        set_enable_true();
                        return;
                    }

                    bool Template_is_open = false;
                    foreach (Document Doc in DocumentManager1)
                    {
                        if (Doc.Name == strTemplatePath)
                        {
                            Template_is_open = true;
                            ThisDrawing = Doc;
                            DocumentManager1.CurrentDocument = ThisDrawing;

                        }

                    }

                    if (Template_is_open == false)
                    {
                        ThisDrawing = DocumentCollectionExtension.Open(DocumentManager1, strTemplatePath, false);

                        Template_is_open = true;
                    }

                    string Scale1 = Get_combobox_viewport_scale_text();

                    if (Functions.IsNumeric(Scale1) == true)
                    {
                        _SGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                    }
                    else
                    {
                        if (Scale1.Contains(":") == true)
                        {
                            Scale1 = Scale1.Replace("1:", "");
                            if (Functions.IsNumeric(Scale1) == true)
                            {
                                _SGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
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
                                _SGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                            }
                        }
                    }



                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                    this.MdiParent.WindowState = FormWindowState.Minimized;

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
                                this.MdiParent.WindowState = FormWindowState.Normal;
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }

                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;

                            Alignment_mdi.Jig_rectangle_viewport_pick_points Jig2 = new Alignment_mdi.Jig_rectangle_viewport_pick_points();
                            Point_res2 = Jig2.StartJig(Point_res1.Value, 1, "\npick top right corner of the plan view");

                            if (Point_res2.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                this.MdiParent.WindowState = FormWindowState.Normal;
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

                            _SGEN_mainform.Band_Separation = Math.Ceiling(3 * (Math.Abs(y2 - y1)) / 10) * 10;


                            if (_SGEN_mainform.Data_Table_regular_bands != null && _SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {

                                for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                {
                                    string CT = _SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"].ToString();
                                    if (_SGEN_mainform.nume_main_vp == CT)
                                    {
                                        _SGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                        _SGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                        _SGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                        _SGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _SGEN_mainform.Vw_scale;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] = _SGEN_mainform.nume_main_vp;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] = Math.Abs(y2 - y1);
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_width"] = Math.Abs(x2 - x1);
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_x"] = (x1 + x2) / 2;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ps_y"] = (y1 + y2) / 2;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_separation"] = _SGEN_mainform.Band_Separation;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] = DBNull.Value;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] = DBNull.Value;
                                        _SGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_twist"] = DBNull.Value;

                                        i = _SGEN_mainform.Data_Table_regular_bands.Rows.Count;


                                    }
                                }

                            }
                            else
                            {

                                if (_SGEN_mainform.Data_Table_regular_bands == null)
                                {
                                    _SGEN_mainform.Data_Table_regular_bands = creeaza_regular_band_data_table_structure();
                                }
                                _SGEN_mainform.Vw_width = Math.Abs(x1 - x2);
                                _SGEN_mainform.Vw_height = Math.Abs(y1 - y2);
                                _SGEN_mainform.Vw_ps_x = (x1 + x2) / 2;
                                _SGEN_mainform.Vw_ps_y = (y1 + y2) / 2;

                                _SGEN_mainform.Data_Table_regular_bands.Rows.Add();

                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = _SGEN_mainform.nume_main_vp;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["Custom_scale"] = _SGEN_mainform.Vw_scale;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_name"] = _SGEN_mainform.nume_main_vp;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_height"] = Math.Abs(y2 - y1);
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_width"] = Math.Abs(x2 - x1);
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_x"] = (x1 + x2) / 2;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ps_y"] = (y1 + y2) / 2;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["band_separation"] = _SGEN_mainform.Band_Separation;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_x"] = DBNull.Value;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_ms_y"] = DBNull.Value;
                                _SGEN_mainform.Data_Table_regular_bands.Rows[_SGEN_mainform.Data_Table_regular_bands.Rows.Count - 1]["viewport_twist"] = DBNull.Value;


                            }


                            #endregion

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

                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_SGEN_mainform.config_file);

                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                    Microsoft.Office.Interop.Excel.Worksheet W_reg = null;

                    if (Workbook1.Worksheets.Count > 1)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet w2 in Workbook1.Worksheets)
                        {
                            if (w2.Name == "Regular_band_data")
                            {
                                W_reg = w2;
                            }
                        }
                    }

                    try
                    {

                        if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                        {
                            if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                            {
                                if (band_name == _SGEN_mainform.nume_main_vp)
                                {
                                    //mainVP
                                    W1.Range["B36"].Value = "True";
                                }
                            }
                        }



                        transfera_regular_band_to_excel(Workbook1);



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
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }
                    creeaza_display_data_table(Creaza_lista_regular_vp_picked());
                    this.MdiParent.WindowState = FormWindowState.Normal;
                }



            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            set_enable_true();
        }

        public string get_template_name_from_text_box()
        {
            return textBox_template_name.Text;
        }

        public string get_textbox_autpot_content()
        {
            if (System.IO.Directory.Exists(textBox_output_folder.Text) == true)
            {
                _SGEN_mainform.output_folder = textBox_output_folder.Text;
                return textBox_output_folder.Text;
            }
            else
            {
                _SGEN_mainform.output_folder = "";
                return "";
            }
        }

        public void transfera_regular_band_to_excel(Microsoft.Office.Interop.Excel.Workbook Workbook1)
        {
            if (_SGEN_mainform.Data_Table_regular_bands != null)
            {
                if (_SGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        if (_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_SGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                            if (bn == _SGEN_mainform.nume_main_vp)
                            {
                                _SGEN_mainform.Data_Table_regular_bands.Rows[i]["Custom_scale"] = _SGEN_mainform.Vw_scale;
                            }

                        }
                    }


                    Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                    foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                    {
                        if (wsh1.Name == "Regular_band_data")
                        {
                            W1 = wsh1;
                        }
                    }

                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W1.Name = "Regular_band_data";

                    }

                    W1.Columns["A:XX"].Delete();
                    W1.Range["A:S"].ColumnWidth = 18;
                    W1.Range["C:K"].ColumnWidth = 2;

                    int maxRows = _SGEN_mainform.Data_Table_regular_bands.Rows.Count;
                    int maxCols = _SGEN_mainform.Data_Table_regular_bands.Columns.Count;

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[maxRows + 1, maxCols]];
                    object[,] values1 = new object[maxRows, maxCols];

                    for (int i = 0; i < maxRows; ++i)
                    {
                        for (int j = 0; j < maxCols; ++j)
                        {
                            if (_SGEN_mainform.Data_Table_regular_bands.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = _SGEN_mainform.Data_Table_regular_bands.Rows[i][j];
                            }
                        }
                    }

                    for (int i = 0; i < _SGEN_mainform.Data_Table_regular_bands.Columns.Count; ++i)
                    {
                        W1.Cells[1, i + 1].value2 = _SGEN_mainform.Data_Table_regular_bands.Columns[i].ColumnName;
                    }

                    range1.Cells.NumberFormat = "@";
                    range1.Value2 = values1;

                    Functions.Color_border_range_inside(range1, 0);

                }
            }

        }

        private void button_align_config_saveall_Click(object sender, EventArgs e)
        {
            current_segment = comboBox_segment_name.Text;

            update_dt_segments(current_segment);
            button_align_config_saveall_boolean(true);
            radioButton_Load_config.Checked = true;
        }

        public void update_dt_segments(string segment1)
        {

            if (_SGEN_mainform.dt_segments != null && _SGEN_mainform.dt_segments.Rows.Count > 0)
            {
                if (segment1 != "")
                {
                    for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                    {
                        if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                        {
                            string segment2 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);

                            if (segment2 == segment1)
                            {

                                if (textBox_template_name.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Template"] = textBox_template_name.Text;
                                }


                                if (textBox_output_folder.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Output Folder"] = textBox_output_folder.Text;
                                }



                                if (textBox_prefix_name.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Prefix File Name"] = textBox_prefix_name.Text;
                                }


                                if (textBox_suffix.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Suffix File Name"] = textBox_suffix.Text;
                                }


                                if (textBox_name_start_number.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Start numbering"] = textBox_name_start_number.Text;
                                }


                                if (textBox_name_increment.Text != "")
                                {
                                    _SGEN_mainform.dt_segments.Rows[i]["Increment"] = textBox_name_increment.Text;
                                }
                            }
                        }
                    }
                }
            }
        }


        public void button_align_config_saveall_boolean(bool Close_dwt)
        {


            if (Functions.Get_if_workbook_is_open_in_Excel(_SGEN_mainform.config_file) == true)
            {
                MessageBox.Show("Please close the " + _SGEN_mainform.config_file + " file");
                return;
            }


            if (Close_dwt == true) close_template();

            if (radioButton_new_config.Checked == true)
            {
                _SGEN_mainform.config_file = "";
            }

            try
            {
                System.Data.DataTable Data_table_config = new System.Data.DataTable();
                Data_table_config.Columns.Add("A", typeof(string));
                Data_table_config.Columns.Add("B", typeof(string));



                for (int i = 0; i <= 45; ++i)
                {
                    Data_table_config.Rows.Add();
                }

                Data_table_config.Rows[0][0] = "Client Name";
                Data_table_config.Rows[0][1] = textBox_client_name.Text;


                Data_table_config.Rows[1][0] = "Project Name";
                Data_table_config.Rows[1][1] = textBox_project_name.Text;

                Data_table_config.Rows[2][0] = "No of Segments";
                Data_table_config.Rows[2][1] = Convert.ToString(_SGEN_mainform.no_of_segments);


                Data_table_config.Rows[3][0] = "Template";
                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[3][1] = get_template_name_from_text_box();

                Data_table_config.Rows[4][0] = "Output folder";

                string out1 = get_output_folder_from_text_box();
                if (out1.Length > 0)
                {
                    if (out1.Substring(out1.Length - 1, 1) != "\\")
                    {
                        out1 = out1 + "\\";
                    }
                }

                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[4][1] = out1;

                Data_table_config.Rows[5][0] = "Prefix File Name";
                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[5][1] = get_prefix_name_from_text_box();

                Data_table_config.Rows[6][0] = "Suffix File Name";
                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[6][1] = get_suffix_name_from_text_box();

                Data_table_config.Rows[7][0] = "Start numbering";
                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[7][1] = get_start_number_from_text_box();

                Data_table_config.Rows[8][0] = "Increment";
                if (_SGEN_mainform.no_of_segments == 0) Data_table_config.Rows[8][1] = get_increment_from_text_box();

                Data_table_config.Rows[9][0] = "Segment Name";



                Data_table_config.Rows[10][0] = "Empty";


                Data_table_config.Rows[11][0] = "Empty";


                Data_table_config.Rows[12][0] = "Empty";
                Data_table_config.Rows[12][1] = "";

                Data_table_config.Rows[13][0] = "Empty";
                Data_table_config.Rows[13][1] = "";


                Data_table_config.Rows[14][0] = "Units";


                if (_SGEN_mainform.units_of_measurement == "m")
                {
                    Data_table_config.Rows[14][1] = "meters";
                }
                else
                {
                    Data_table_config.Rows[14][1] = "feet";
                }

                Data_table_config.Rows[15][0] = "Project database folder location";


                Data_table_config.Rows[15][1] = _SGEN_mainform.project_main_folder;

                Data_table_config.Rows[16][0] = "Empty";
                Data_table_config.Rows[16][1] = "";



                Data_table_config.Rows[17][0] = "Empty";
                Data_table_config.Rows[17][1] = "";

                Data_table_config.Rows[18][0] = "Empty";


                Data_table_config.Rows[19][0] = "Empty";


                Data_table_config.Rows[20][0] = "Empty";



                Data_table_config.Rows[21][0] = "Empty";



                Data_table_config.Rows[22][0] = "Empty";


                Data_table_config.Rows[23][0] = "Empty";
                Data_table_config.Rows[23][1] = "";

                Data_table_config.Rows[24][0] = "Empty";
                Data_table_config.Rows[24][1] = "";

                Data_table_config.Rows[25][0] = "Empty";
                Data_table_config.Rows[25][1] = "";

                Data_table_config.Rows[26][0] = "Empty";
                Data_table_config.Rows[26][1] = "";

                Data_table_config.Rows[27][0] = "Empty";
                Data_table_config.Rows[27][1] = "";

                Data_table_config.Rows[28][0] = "Empty";
                Data_table_config.Rows[28][1] = "";

                Data_table_config.Rows[29][0] = "Empty";
                Data_table_config.Rows[29][1] = "";

                Data_table_config.Rows[30][0] = "Empty";
                Data_table_config.Rows[30][1] = "";

                Data_table_config.Rows[31][0] = "Empty";
                Data_table_config.Rows[31][1] = "";

                Data_table_config.Rows[32][0] = "Empty";
                Data_table_config.Rows[32][1] = "";

                Data_table_config.Rows[33][0] = "Empty";
                Data_table_config.Rows[33][1] = "";

                Data_table_config.Rows[34][0] = "Empty";
                Data_table_config.Rows[34][1] = "";

                Data_table_config.Rows[35][0] = "Main viewport picked";

                if (_SGEN_mainform.Vw_height > 0 && _SGEN_mainform.Vw_width > 0)
                {
                    Data_table_config.Rows[35][1] = "TRUE";

                }
                else
                {
                    Data_table_config.Rows[35][1] = "FALSE";
                }






                string Scale1 = Get_combobox_viewport_scale_text();

                if (Functions.IsNumeric(Scale1) == true)
                {
                    _SGEN_mainform.Vw_scale = Convert.ToDouble(Scale1);
                }
                else
                {
                    if (Scale1.Contains(":") == true)
                    {
                        Scale1 = Scale1.Replace("1:", "");
                        if (Functions.IsNumeric(Scale1) == true)
                        {
                            _SGEN_mainform.Vw_scale = 1000 / Convert.ToDouble(Scale1);
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
                            _SGEN_mainform.Vw_scale = 1 / Convert.ToDouble(Scale1);
                        }
                    }
                }


                if (System.IO.File.Exists(_SGEN_mainform.config_file) == true)
                {
                    update_config_file(Data_table_config);
                }
                else
                {
                    save_new_config_file(Data_table_config);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void close_template()
        {

            try
            {



                string strTemplatePath = get_template_name_from_text_box();

                DocumentCollection DocumentManager1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

                // Ag.WindowState = FormWindowState.Minimized;
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

            this.MdiParent.WindowState = FormWindowState.Normal;


        }
        public string get_output_folder_from_text_box()
        {
            return textBox_output_folder.Text;
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


        private void save_new_config_file(System.Data.DataTable Data_table_config)
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet W1 = Workbook1.Worksheets[1];
                W1.Name = "main_cfg";


                try
                {



                    SaveFileDialog Save_dlg = new SaveFileDialog();
                    Save_dlg.Filter = "Excel file|*.xlsx";


                    if (Save_dlg.ShowDialog() == DialogResult.OK)
                    {

                        if (System.IO.File.Exists(Save_dlg.FileName) == false)
                        {
                            _SGEN_mainform.config_file = Save_dlg.FileName;
                            _SGEN_mainform.project_main_folder = System.IO.Path.GetDirectoryName(_SGEN_mainform.config_file);
                            Data_table_config.Rows[15][1] = _SGEN_mainform.project_main_folder;

                            W1.Cells.NumberFormat = "@";
                            W1.Range["C4:" + Functions.get_excel_column_letter(100) + "10"].ClearContents();


                            int maxRows = Data_table_config.Rows.Count, maxCols = Data_table_config.Columns.Count;
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                            object[,] values = new object[maxRows, maxCols];
                            for (int row = 0; row < maxRows; row++)
                            {
                                for (int col = 0; col < maxCols; col++)
                                {
                                    if (Data_table_config.Rows[row][col] != DBNull.Value)
                                    {
                                        values[row, col] = Data_table_config.Rows[row][col];
                                    }
                                }
                            }
                            range1.Value2 = values;


                            if (_SGEN_mainform.no_of_segments > 0)
                            {

                                W1.Range["B4:B9"].ClearContents();
                                string end_cell = Functions.get_excel_column_letter(_SGEN_mainform.no_of_segments + 2) + "10";

                                Microsoft.Office.Interop.Excel.Range range2 = W1.Range["C4:" + end_cell];

                                object[,] values2 = new object[_SGEN_mainform.dt_segments.Columns.Count, _SGEN_mainform.dt_segments.Rows.Count];
                                for (int row = 0; row < _SGEN_mainform.dt_segments.Rows.Count; row++)
                                {
                                    for (int col = 0; col < _SGEN_mainform.dt_segments.Columns.Count; col++)
                                    {
                                        if (_SGEN_mainform.dt_segments.Rows[row][col] != DBNull.Value)
                                        {
                                            values2[col, row] = _SGEN_mainform.dt_segments.Rows[row][col];
                                        }
                                    }
                                }
                                range2.Value2 = values2;


                            }

                            range1.Columns.AutoFit();

                            transfera_regular_band_to_excel(Workbook1);





                            Workbook1.SaveAs(_SGEN_mainform.config_file);
                        }
                        else
                        {
                            MessageBox.Show("File exists\r\nOperation aborted\r\nSpecify another name.....");
                            return;
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


        private void update_config_file(System.Data.DataTable Data_table_config)
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

                if (Excel1.Workbooks.Count == 0) Excel1.Visible = true;



                Excel1.Visible = true;



                if (System.IO.File.Exists(_SGEN_mainform.config_file) == true)
                {

                    Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(_SGEN_mainform.config_file);
                    Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                    W1.Name = "main_cfg";
                    try
                    {

                        W1.Range["C4:" + Functions.get_excel_column_letter(100) + "10"].ClearContents();

                        W1.Cells.NumberFormat = "@";
                        int maxRows = Data_table_config.Rows.Count;
                        int maxCols = Data_table_config.Columns.Count;
                        Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[1, 1], W1.Cells[maxRows, maxCols]];

                        object[,] values = new object[maxRows, maxCols];
                        for (int row = 0; row < maxRows; row++)
                        {
                            for (int col = 0; col < maxCols; col++)
                            {
                                if (Data_table_config.Rows[row][col] != DBNull.Value)
                                {
                                    values[row, col] = Data_table_config.Rows[row][col];
                                }
                            }
                        }
                        range1.Value2 = values;

                        if (_SGEN_mainform.no_of_segments > 0)
                        {

                            W1.Range["B4:B9"].ClearContents();

                            string end_cell = Functions.get_excel_column_letter(_SGEN_mainform.no_of_segments + 2) + "10";

                            Microsoft.Office.Interop.Excel.Range range2 = W1.Range["C4:" + end_cell];

                            object[,] values2 = new object[_SGEN_mainform.dt_segments.Columns.Count, _SGEN_mainform.dt_segments.Rows.Count];
                            for (int row = 0; row < _SGEN_mainform.dt_segments.Rows.Count; row++)
                            {
                                for (int col = 0; col < _SGEN_mainform.dt_segments.Columns.Count; col++)
                                {
                                    if (_SGEN_mainform.dt_segments.Rows[row][col] != DBNull.Value)
                                    {
                                        values2[col, row] = _SGEN_mainform.dt_segments.Rows[row][col];
                                    }
                                }
                            }
                            range2.Value2 = values2;


                        }



                        range1.Columns.AutoFit();

                        transfera_regular_band_to_excel(Workbook1);


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
                        if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                        if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                    }

                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void radioButton_new_config_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_new_config.Checked == true)
            {
                _SGEN_mainform.config_file = "";
                _SGEN_mainform.project_main_folder = "";
            }
        }

        public string Get_client_name()
        {
            return textBox_client_name.Text;
        }
        public string Get_project_name()
        {
            return textBox_project_name.Text;
        }

        private void button_show_segment_list_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi.AGEN_segments_form)
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
                Alignment_mdi.AGEN_segments_form forma2 = new Alignment_mdi.AGEN_segments_form();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
        }

        public string get_combobox_segment_name_value()
        {
            return comboBox_segment_name.Text;
        }

        public bool check_combobox_segment_is_first_one()
        {
            if (comboBox_segment_name.SelectedIndex == 0)
            {
                return true;
            }
            return false;
        }

        private void comboBox_segment_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            update_dt_segments(current_segment);

            string segment1 = comboBox_segment_name.Text;

            populate_controls_on_page(segment1);

            current_segment = comboBox_segment_name.Text;

        }

        public void populate_controls_on_page(string segment1)
        {
            if (segment1 != "")
            {
                current_segment = segment1;

                for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                {
                    if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                    {
                        string segment2 = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]);

                        if (segment2 == segment1)
                        {
                            if (_SGEN_mainform.dt_segments.Rows[i]["Template"] != DBNull.Value)
                            {
                                textBox_template_name.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Template"]);
                            }
                            else
                            {
                                textBox_template_name.Text = "";
                            }

                            if (_SGEN_mainform.dt_segments.Rows[i]["Output Folder"] != DBNull.Value)
                            {
                                textBox_output_folder.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Output Folder"]);
                            }
                            else
                            {
                                textBox_output_folder.Text = "";
                            }


                            if (_SGEN_mainform.dt_segments.Rows[i]["Prefix File Name"] != DBNull.Value)
                            {
                                textBox_prefix_name.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Prefix File Name"]);
                            }
                            else
                            {
                                textBox_prefix_name.Text = "";
                            }


                            if (_SGEN_mainform.dt_segments.Rows[i]["Suffix File Name"] != DBNull.Value)
                            {
                                textBox_suffix.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Suffix File Name"]);
                            }
                            else
                            {
                                textBox_suffix.Text = "";
                            }

                            if (_SGEN_mainform.dt_segments.Rows[i]["Start numbering"] != DBNull.Value)
                            {
                                textBox_name_start_number.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Start numbering"]);
                            }
                            else
                            {
                                textBox_name_start_number.Text = "";
                            }

                            if (_SGEN_mainform.dt_segments.Rows[i]["Increment"] != DBNull.Value)
                            {
                                textBox_name_increment.Text = Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Increment"]);
                            }
                            else
                            {
                                textBox_name_increment.Text = "";
                            }

                            label_sheet_naming.Text = "Sheet Setup " + segment1;

                            _SGEN_mainform.tpage_sheetindex.Build_sheet_index_dt_from_excel();
                        }

                    }
                }
            }
        }


        public void Fill_combobox_segments()
        {
            comboBox_segment_name.Items.Clear();
            if (_SGEN_mainform.dt_segments != null && _SGEN_mainform.dt_segments.Rows.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _SGEN_mainform.dt_segments.Rows.Count; ++i)
                    {
                        if (_SGEN_mainform.dt_segments.Rows[i]["Segment Name"] != DBNull.Value)
                        {
                            comboBox_segment_name.Items.Add(Convert.ToString(_SGEN_mainform.dt_segments.Rows[i]["Segment Name"]));
                        }

                    }
                    if (comboBox_segment_name.Items.Count > 0) comboBox_segment_name.SelectedIndex = 0;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        public string get_textBox_client_name()
        {
            return textBox_client_name.Text;
        }
        public string get_textBox_project_name()
        {
            return textBox_project_name.Text;
        }
        public string get_textBox_prefix_name()
        {
            return textBox_prefix_name.Text;
        }
    }
}
