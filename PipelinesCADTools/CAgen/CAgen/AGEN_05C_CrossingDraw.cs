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
    public partial class AGEN_Crossing_draw : Form
    {
        string block_pi_name = "PI";
        string block_xing_name = "XING";
        string block_prop_name = "PROP";
        string col_xing_block = "Block Name";
        string col_xing_sta = "Attrib Sta";
        string col_xing_descr = "Attrib Desc";
        string atr_sta = "STA";
        string atr_descr = "DESCR";
        string atr_pi_sta = "STA";
        string atr_pi_descr = "DESCR";
        string atr_prop_sta = "STA";
        string atr_prop_descr = "DESCR";
        string col_visibility = "Visibility";


        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_crossing_band);
            lista_butoane.Add(button_show_crossing_band_settings);
            lista_butoane.Add(comboBox_crossing_pi_textstyle);
            lista_butoane.Add(comboBox_crossing_textstyle);
            lista_butoane.Add(comboBox_end);
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(checkBox_display_station);
            lista_butoane.Add(checkBox_draw_angle_symbol);
            lista_butoane.Add(checkBox_include_property_lines);
            lista_butoane.Add(checkBox_pi_underline);
            lista_butoane.Add(checkBox_split_station);
            lista_butoane.Add(textBox_crossing_text_rotation);
            lista_butoane.Add(textBox_pi_min_angle);
            lista_butoane.Add(textBox_pi_prefix);
            lista_butoane.Add(textBox_rounding_decimal_degrees);
            lista_butoane.Add(textBox_station_prefix);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {

                bt1.Enabled = false;

            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_crossing_band);
            lista_butoane.Add(button_show_crossing_band_settings);
            lista_butoane.Add(comboBox_crossing_pi_textstyle);
            lista_butoane.Add(comboBox_crossing_textstyle);
            lista_butoane.Add(comboBox_end);
            lista_butoane.Add(comboBox_start);
            lista_butoane.Add(checkBox_display_station);
            lista_butoane.Add(checkBox_draw_angle_symbol);
            lista_butoane.Add(checkBox_include_property_lines);
            lista_butoane.Add(checkBox_pi_underline);
            lista_butoane.Add(checkBox_split_station);
            lista_butoane.Add(textBox_crossing_text_rotation);
            lista_butoane.Add(textBox_pi_min_angle);
            lista_butoane.Add(textBox_pi_prefix);
            lista_butoane.Add(textBox_rounding_decimal_degrees);
            lista_butoane.Add(textBox_station_prefix);
            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public AGEN_Crossing_draw()
        {
            InitializeComponent();
        }

        private void button_refresh_pi_textstyle_Click(object sender, EventArgs e)
        {

            try
            {
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Functions.Incarca_existing_textstyles_to_combobox(comboBox_crossing_textstyle);
                        Functions.Incarca_existing_textstyles_to_combobox(comboBox_crossing_pi_textstyle);
                        if (comboBox_crossing_textstyle.Items.Contains("Agen_Text_Crossing") == true)
                        {
                            comboBox_crossing_textstyle.SelectedIndex = comboBox_crossing_textstyle.Items.IndexOf("Agen_Text_Crossing");
                        }
                        else
                        {
                            comboBox_crossing_textstyle.SelectedIndex = 0;
                        }

                        if (comboBox_crossing_pi_textstyle.Items.Contains("Agen_Text_PI") == true)
                        {
                            comboBox_crossing_pi_textstyle.SelectedIndex = comboBox_crossing_pi_textstyle.Items.IndexOf("Agen_Text_PI");
                        }
                        else
                        {
                            comboBox_crossing_pi_textstyle.SelectedIndex = 0;
                        }

                        Trans1.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            set_enable_true();


        }

        public string get_comboBox_crossing_textstyle()
        {
            return comboBox_crossing_textstyle.Text;
        }

        public void set_comboBox_crossing_textstyle(string txt)
        {
            if (txt != "")
            {
                if (comboBox_crossing_textstyle.Items.Contains(txt) == false)
                {
                    comboBox_crossing_textstyle.Items.Add(txt);

                }
                comboBox_crossing_textstyle.SelectedIndex = comboBox_crossing_textstyle.Items.IndexOf(txt);
            }
        }
        public string get_comboBox_crossing_pi_textstyle()
        {
            return comboBox_crossing_pi_textstyle.Text;
        }

        public void set_comboBox_crossing_pi_textstyle(string txt)
        {
            if (txt != "")
            {
                if (comboBox_crossing_pi_textstyle.Items.Contains(txt) == false)
                {
                    comboBox_crossing_pi_textstyle.Items.Add(txt);

                }
                comboBox_crossing_pi_textstyle.SelectedIndex = comboBox_crossing_pi_textstyle.Items.IndexOf(txt);
            }
        }



        public string get_textBox_crossing_text_rotation()
        {
            return textBox_crossing_text_rotation.Text;
        }

        public void set_textBox_crossing_text_rotation(string txt)
        {
            textBox_crossing_text_rotation.Text = txt;
        }


        public string get_textBox_pi_min_angle()
        {
            return textBox_pi_min_angle.Text;
        }

        public void set_textBox_pi_min_angle(string txt)
        {
            textBox_pi_min_angle.Text = txt;
        }

        public string get_textBox_pi_prefix()
        {
            return textBox_pi_prefix.Text;
        }

        public void set_textBox_pi_prefix(string txt)
        {
            textBox_pi_prefix.Text = txt;
        }

        public bool get_checkBox_pi_underline_value()
        {
            return checkBox_pi_underline.Checked;
        }

        public void set_checkBox_pi_underline_value(bool val)
        {
            checkBox_pi_underline.Checked = val;
        }

        public bool get_checkBox_display_station_value()
        {
            return checkBox_display_station.Checked;
        }

        public bool get_checkBox_include_property_lines()
        {
            return checkBox_include_property_lines.Checked;
        }

        public double get_text_box_overwrite_value()
        {
            double th = -1;

            if (checkBox_overwrite_text_height.Checked == true)
            {
                if (Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
                {
                    th = Convert.ToDouble(textBox_overwrite_text_height.Text);
                }
            }

            return th;
        }

        public void set_checkBox_display_station(bool val)
        {
            checkBox_display_station.Checked = val;
        }

        public string get_textBox_station_prefix()
        {
            return textBox_station_prefix.Text;
        }

        public void set_textBox_station_prefix(string txt)
        {
            textBox_station_prefix.Text = txt;
        }

        public void set_textBox_rounding_decimal_degrees(string val)
        {
            textBox_rounding_decimal_degrees.Text = val;
        }

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }
        public string get_textBox_rounding_decimal_degrees()
        {
            return textBox_rounding_decimal_degrees.Text;
        }

        public bool get_checkBox_split_station_value()
        {
            return checkBox_split_station.Checked;
        }

        public void set_checkBox_split_station_value(bool val)
        {
            checkBox_split_station.Checked = val;
        }

        public void set_checkBox_include_property_lines(bool chck)
        {
            checkBox_include_property_lines.Checked = chck;
        }
        public void set_checkBox_draw_angle_symbol_value(bool chck)
        {
            checkBox_draw_angle_symbol.Checked = chck;
        }

        public bool get_checkBox_draw_angle_symbol_value()
        {
            return checkBox_draw_angle_symbol.Checked;
        }
        public bool get_checkBox_use_blocks_value()
        {
            return checkBox_use_blocks.Checked;
        }

        public void clear_combobox()
        {
            comboBox_sheet_index_tabs.Items.Clear();
            comboBox_crossings_tabs.Items.Clear();
        }

        public void set_checkBox_use_blocks(bool chck)
        {
            checkBox_use_blocks.Checked = chck;
        }

        public void write_crossing_settings_to_excel(bool ExcelVisible, string File1, System.Data.DataTable dt_dwg_data)
        {


            string ts1 = get_comboBox_crossing_textstyle();
            string ts2 = get_comboBox_crossing_pi_textstyle();
            string ts3 = get_textBox_crossing_text_rotation();
            string ts4 = get_textBox_pi_min_angle();
            string ts5 = get_textBox_pi_prefix();
            string ts6 = get_checkBox_pi_underline_value().ToString();
            string ts7 = get_checkBox_display_station_value().ToString();
            string ts8 = get_textBox_station_prefix();
            string ts9 = get_textBox_rounding_decimal_degrees();
            string ts10 = get_checkBox_split_station_value().ToString();
            string ts11 = get_checkBox_draw_angle_symbol_value().ToString();
            string ts12 = get_checkBox_include_property_lines().ToString();
            string ts13 = get_checkBox_use_blocks_value().ToString();
            double th = get_text_box_overwrite_value();

            if (ts1 != "" || ts2 != "" || ts3 != "" || ts4 != "" || ts5 != "" || ts6 != "" || ts7 != "" || ts8 != "" || ts9 != "" || ts10 != "" || ts11 != "" || ts12 != "" || ts13 != "" || th > 0)
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
                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                {
                    if (wsh1.Name == "Crossing_data_config")
                    {
                        W1 = wsh1;
                    }
                }

                if (W1 == null)
                {
                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W1.Name = "Crossing_data_config";
                }

                try
                {
                    int NrR = 17;
                    int NrC = 2;

                    object[,] values = new object[NrR, NrC];
                    values[0, 0] = "Crossing Text style";
                    values[0, 1] = ts1;
                    values[1, 0] = "PI Text style";
                    values[1, 1] = ts2;
                    values[2, 0] = "Text rotation";
                    values[2, 1] = ts3;
                    values[3, 0] = "PI minimum angle";
                    values[3, 1] = ts4;
                    values[4, 0] = "PI prefix";
                    values[4, 1] = ts5;
                    values[5, 0] = "underline PI";
                    values[5, 1] = ts6;
                    values[6, 0] = "Display station";
                    values[6, 1] = ts7;
                    values[7, 0] = "Station prefix";
                    values[7, 1] = ts8;
                    values[8, 0] = "Deflection rounding [decimal degrees]";
                    values[8, 1] = ts9;
                    values[9, 0] = "Split station and description";
                    values[9, 1] = ts10;
                    values[10, 0] = "Draw angle symbol";
                    values[10, 1] = ts11;
                    values[11, 0] = "Include Property lines";
                    values[11, 1] = ts12;
                    values[16, 0] = "Use blocks";
                    values[16, 1] = ts13;
                    values[12, 0] = "Delta Y Station";
                    values[12, 1] = _AGEN_mainform.XingDeltay1.ToString();
                    values[13, 0] = "Delta Y Symbol";
                    values[13, 1] = _AGEN_mainform.XingDeltay2.ToString();
                    values[14, 0] = "Delta Y Description";
                    values[14, 1] = _AGEN_mainform.XingDeltay3.ToString();
                    values[15, 0] = "Overwrite Text Height Value";
                    if (th > 0)
                    {
                        values[15, 1] = Math.Round(th, 1);
                    }
                    else
                    {
                        values[15, 1] = null;
                        th = -1;
                    }

                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B17"];
                    W1.Range["A:A"].ColumnWidth = 35;
                    W1.Range["B:B"].ColumnWidth = 9;
                    range1.Cells.NumberFormat = "General";
                    range1.Value2 = values;
                    Functions.Color_border_range_inside(range1, 0);

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


        private void button_show_crossing_band_settings_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Hide();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_sheetindex.Hide();
            _AGEN_mainform.tpage_layer_alias.Hide();

            _AGEN_mainform.tpage_crossing_draw.Hide();
            _AGEN_mainform.tpage_profilescan.Hide();
            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_crossing_scan.Show();
        }

        public System.Data.DataTable Load_existing_crossing(string file_xing, string sheetname = "", bool add_size_to_dt2 = false)
        {
            System.Data.DataTable dt2 = new System.Data.DataTable();
            if (System.IO.File.Exists(file_xing) == false)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nthe crosings data file does not exist");
                dt2 = Functions.Creaza_crossing_datatable_structure();
                return dt2;
            }


            try
            {
                bool excel_is_opened = false;
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == file_xing.ToLower())
                        {
                            Workbook1 = Workbook2;
                            if (sheetname == "")
                            {
                                W1 = Workbook1.Worksheets[1];
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                                {
                                    if (W2.Name.ToLower() == sheetname.ToLower())
                                    {
                                        W1 = W2;
                                    }
                                }
                                if (W1 == null)
                                {
                                    W1 = Workbook1.Worksheets[1];
                                }
                            }
                            excel_is_opened = true;
                        }

                    }


                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }
                if (W1 == null)
                {
                    if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                    Workbook1 = Excel1.Workbooks.Open(file_xing);
                    foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                    {
                        if (W2.Name.ToLower() == sheetname.ToLower())
                        {
                            W1 = W2;
                        }
                    }
                    if (W1 == null)
                    {
                        W1 = Workbook1.Worksheets[1];
                    }
                }


                try
                {
                   
                 


                    dt2 = Functions.Build_Data_table_crossings_from_excel(W1, _AGEN_mainform.Start_row_crossing + 1, add_size_to_dt2);

                    string val1 = W1.Range["D1"].Value2;
                    if (val1 != "")
                    {
                        block_pi_name = val1;
                    }
                    string val2 = W1.Range["D2"].Value2;
                    if (val2 != "")
                    {
                        atr_pi_sta = val2;
                    }

                    string val3 = W1.Range["D3"].Value2;
                    if (val3 != "")
                    {
                        atr_pi_descr = val3;
                    }

                    string val4 = W1.Range["D4"].Value2;
                    if (val4 != "")
                    {
                        block_prop_name = val4;
                    }
                    string val5 = W1.Range["D5"].Value2;
                    if (val5 != "")
                    {
                        atr_prop_sta = val5;
                    }

                    string val6 = W1.Range["D6"].Value2;
                    if (val6 != "")
                    {
                        atr_prop_descr = val6;
                    }


                    if (excel_is_opened == false)
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



        private void set_checkbox_overwrite(string texth, bool chck)
        {
            checkBox_overwrite_text_height.Checked = chck;
            textBox_overwrite_text_height.Text = texth;
        }

        public void write_crossing_settings_to_controls(string crossingtextstyle, string pitextstyle, string textrotation,
            string piminangle, string piprefix, bool underline, bool displaysta, string staprefix,
            string defl_rounding, bool split, bool draw_angle, bool incl_property, bool useblk, double th)
        {
            try
            {
                set_comboBox_crossing_textstyle(crossingtextstyle);
                set_comboBox_crossing_pi_textstyle(pitextstyle);
                set_textBox_crossing_text_rotation(textrotation);
                set_textBox_pi_min_angle(piminangle);
                set_textBox_pi_prefix(piprefix);
                set_checkBox_pi_underline_value(underline);
                set_checkBox_display_station(displaysta);
                set_textBox_station_prefix(staprefix);
                set_textBox_rounding_decimal_degrees(defl_rounding);
                set_checkBox_split_station_value(split);
                set_checkBox_draw_angle_symbol_value(draw_angle);
                set_checkBox_include_property_lines(incl_property);
                set_checkBox_use_blocks(useblk);
                if (th > 0)
                {
                    set_checkbox_overwrite(Functions.Get_String_Rounded(th, 1), true);
                }
                else
                {
                    set_checkbox_overwrite("", false);
                }

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
        }



        public void add_measure_column_to_dt_si(System.Data.DataTable dt1, Polyline poly2d, Polyline3d poly3d)
        {
            if (dt1.Columns.Contains("meas1") == false)
            {
                dt1.Columns.Add("meas1", typeof(double));
            }

            if (dt1.Columns.Contains("meas2") == false)
            {
                dt1.Columns.Add("meas2", typeof(double));
            }

            for (int i = dt1.Rows.Count - 1; i >= 0; --i)
            {
                if (dt1.Rows[i]["X_Beg"] != DBNull.Value && dt1.Rows[i]["Y_Beg"] != DBNull.Value && dt1.Rows[i]["X_End"] != DBNull.Value && dt1.Rows[i]["Y_End"] != DBNull.Value)
                {
                    double xm1 = Convert.ToDouble(dt1.Rows[i]["X_Beg"]);
                    double ym1 = Convert.ToDouble(dt1.Rows[i]["Y_Beg"]);
                    double xm2 = Convert.ToDouble(dt1.Rows[i]["X_End"]);
                    double ym2 = Convert.ToDouble(dt1.Rows[i]["Y_End"]);
                    Point3d pt_2d_m1 = poly2d.GetClosestPointTo(new Point3d(xm1, ym1, poly2d.Elevation), Vector3d.ZAxis, false);
                    Point3d pt_2d_m2 = poly2d.GetClosestPointTo(new Point3d(xm2, ym2, poly2d.Elevation), Vector3d.ZAxis, false);
                    double param1 = poly2d.GetParameterAtPoint(pt_2d_m1);
                    double param2 = poly2d.GetParameterAtPoint(pt_2d_m2);
                    if (param1 > poly3d.EndParam) param1 = poly3d.EndParam;
                    if (param2 > poly3d.EndParam) param2 = poly3d.EndParam;
                    double M1 = poly3d.GetDistanceAtParameter(param1);
                    double M2 = poly3d.GetDistanceAtParameter(param2);
                    dt1.Rows[i]["meas1"] = M1;
                    dt1.Rows[i]["meas2"] = M2;
                }
                else
                {
                    dt1.Rows[i].Delete();
                }
            }
        }

        public int populate_extra_columns_on_crossing_table(System.Data.DataTable dt_crossing, System.Data.DataTable dt_si, double sta)
        {
            int idx = -1;
            for (int i = 0; i < dt_si.Rows.Count; ++i)
            {
                if (
                     dt_si.Rows[i]["StaBeg"] != DBNull.Value &&
                    dt_si.Rows[i]["StaEnd"] != DBNull.Value &&
                    dt_si.Rows[i][_AGEN_mainform.Col_dwg_name] != DBNull.Value
                    )
                {
                    string dwg_name = Convert.ToString(dt_si.Rows[i][_AGEN_mainform.Col_dwg_name]);
                    double M1 = Convert.ToDouble(dt_si.Rows[i]["StaBeg"]);
                    double M2 = Convert.ToDouble(dt_si.Rows[i]["StaEnd"]);
                    if (dt_si.Columns.Contains("meas1") == true && dt_si.Columns.Contains("meas2") == true)
                    {
                        M1 = Convert.ToDouble(dt_si.Rows[i]["meas1"]);
                        M2 = Convert.ToDouble(dt_si.Rows[i]["meas2"]);
                    }


                    if (Math.Round(sta, _AGEN_mainform.round1 + 2) >= Math.Round(M1, _AGEN_mainform.round1 + 2) && Math.Round(sta, _AGEN_mainform.round1 + 2) <= Math.Round(M2, _AGEN_mainform.round1 + 2))
                    {
                        idx = i;
                        dt_crossing.Rows[dt_crossing.Rows.Count - 1]["index1"] = i;
                        dt_crossing.Rows[dt_crossing.Rows.Count - 1]["dwg"] = dwg_name;
                        dt_crossing.Rows[dt_crossing.Rows.Count - 1]["DispXing"] = "YES";

                        i = dt_si.Rows.Count;
                    }
                }
            }
            return idx;
        }

        private void button_draw_crossing_band_Click(object sender, EventArgs e)
        {
            string lnp = "Agen_no_plot_cross";





            if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
            {
                MessageBox.Show("no config file loaded\r\nOperation aborted");
                return;
            }

            string cfg1 = System.IO.Path.GetFileName(_AGEN_mainform.config_path);
            if (Functions.Get_if_workbook_is_open_in_Excel(cfg1) == true)
            {
                MessageBox.Show("Please close the " + cfg1 + " file");
                return;
            }


            _AGEN_mainform.tpage_processing.Show();

            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }




                string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;

                if (checkBox_include_property_lines.Checked == true)
                {


                    if (System.IO.File.Exists(fisier_prop) == false)
                    {
                        _AGEN_mainform.tpage_processing.Hide();
                        set_enable_true();
                        MessageBox.Show("the property data file does not exist");
                        return;
                    }

                }


                string fisier_cs = ProjF + _AGEN_mainform.crossing_excel_name;


                _AGEN_mainform.Data_Table_crossings = Load_existing_crossing(fisier_cs, comboBox_crossings_tabs.Text);

                if (checkBox_include_property_lines.Checked == true)
                {
                    _AGEN_mainform.Data_Table_property = _AGEN_mainform.tpage_setup.Load_existing_property(fisier_prop);
                }


                Functions.Load_entities_records_from_config_file(_AGEN_mainform.config_path);

            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            if (comboBox_sheet_index_tabs.Text != "")
            {
                _AGEN_mainform.tpage_setup.Build_sheet_index_dt_from_excel(comboBox_sheet_index_tabs.Text);
            }

            System.Data.DataTable dt_si = _AGEN_mainform.dt_sheet_index.Copy();


            if (dt_si.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the sheet index file does not have any data");
                return;
            }

            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            if (checkBox_include_property_lines.Checked == true)
            {
                if (_AGEN_mainform.Data_Table_property.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    MessageBox.Show("the property file does not have any data\r\n and you selected to append property lines\r\noperation aborted");
                    return;
                }
            }

            System.Data.DataTable dt_for_cfg = Functions.creaza_crossing_block_table_record_structure();
            int debug_i = 0;
            try
            {


                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        Polyline3d poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);


                        int lr = 1;
                        if (_AGEN_mainform.Left_to_Right == false)
                        {
                            lr = -1;
                        }
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;

                        ObjectId Txt_style_id_crossing = Functions.Get_textstyle_id(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_textstyle());
                        ObjectId Txt_style_id_pi_crossing = Functions.Get_textstyle_id(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_pi_textstyle());

                        ObjectId Txt_style_id_standard = Functions.Get_textstyle_id("Standard");


                        if (Txt_style_id_crossing != null)
                        {
                            TextStyleTableRecord TextStyle1 = Trans1.GetObject(Txt_style_id_crossing, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                            TextStyleTableRecord TextStyle2 = Trans1.GetObject(Txt_style_id_pi_crossing, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as TextStyleTableRecord;
                            if (TextStyle1 != null && TextStyle2 != null)
                            {
                                double TextH_crossing = TextStyle1.TextSize;
                                double TextH_pi_crossing = TextStyle2.TextSize;

                                if (checkBox_overwrite_text_height.Checked == true)
                                {
                                    if (Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
                                    {
                                        TextH_crossing = Math.Abs(Convert.ToDouble(textBox_overwrite_text_height.Text));
                                        TextH_pi_crossing = Math.Abs(Convert.ToDouble(textBox_overwrite_text_height.Text));
                                    }
                                }

                                if (checkBox_use_blocks.Checked == false && (TextH_crossing == 0 || TextH_pi_crossing == 0))
                                {
                                    MessageBox.Show("no text height specified");
                                    set_enable_true();
                                    return;
                                }

                                _AGEN_mainform.XingDeltay1 = TextH_crossing;
                                if (TextH_crossing == 0) TextH_crossing = 2.5;
                                if (TextH_pi_crossing == 0) TextH_pi_crossing = 2.5;


                                double min_dist = 2 * TextH_crossing;


                                double min_ang = 0;
                                if (Functions.IsNumeric(_AGEN_mainform.tpage_crossing_draw.get_textBox_pi_min_angle()) == true)
                                {
                                    min_ang = Convert.ToDouble(_AGEN_mainform.tpage_crossing_draw.get_textBox_pi_min_angle());
                                }

                                if (checkBox_use_blocks.Checked == false && (TextH_crossing <= 0 || TextH_pi_crossing <= 0))
                                {
                                    MessageBox.Show("The textstyle you selected does not have a set height. \r\nOperation aborted");
                                    _AGEN_mainform.tpage_processing.Hide();
                                    _AGEN_mainform.tpage_blank.Hide();
                                    _AGEN_mainform.tpage_setup.Hide();
                                    _AGEN_mainform.tpage_viewport_settings.Hide();
                                    _AGEN_mainform.tpage_tblk_attrib.Hide();
                                    _AGEN_mainform.tpage_sheetindex.Hide();
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

                                    _AGEN_mainform.tpage_layer_alias.Show();
                                    set_enable_true();
                                    return;
                                }

                                double TextR = Math.PI / 2;
                                if (Functions.IsNumeric(_AGEN_mainform.tpage_crossing_draw.get_textBox_crossing_text_rotation()) == true)
                                {
                                    TextR = Convert.ToDouble(_AGEN_mainform.tpage_crossing_draw.get_textBox_crossing_text_rotation()) * Math.PI / 180;
                                }
                                double Wfactor = 1;

                                Wfactor = Functions.Get_text_width_factor_from_textstyle(_AGEN_mainform.tpage_crossing_draw.get_comboBox_crossing_textstyle());

                                add_measure_column_to_dt_si(dt_si, poly2d, poly3d);

                                #region lista_bands
                                List<int> lista_bands_for_generation = new List<int>();

                                if (comboBox_start.Text != "" & comboBox_end.Text != "")
                                {
                                    lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg(comboBox_start.Text, comboBox_end.Text);
                                }
                                else
                                {
                                    lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg("", "");
                                }

                                #endregion

                                #region build data table crossing for drafting
                                System.Data.DataTable dt_crossing = Functions.Creaza_crossing_datatable_structure();
                                dt_crossing.Columns.Add("index1", typeof(int));
                                dt_crossing.Columns.Add("dwg", typeof(string));
                                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                                {
                                    #region build from crossings.xls

                                    for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                    {
                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"] != DBNull.Value)
                                        {
                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"].ToString().ToUpper() == "YES" || _AGEN_mainform.Data_Table_crossings.Rows[i]["DispXing"].ToString().ToUpper() == "TRUE")
                                            {
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value && _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value
                                                    && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x])) == true
                                                    && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y])) == true
                                                    && _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.col_desc] != DBNull.Value
                                                    )

                                                {
                                                    dt_crossing.ImportRow(_AGEN_mainform.Data_Table_crossings.Rows[i]);
                                                    double x = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x]);
                                                    double y = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y]);


                                                    Point3d pt_on_line = poly2d.GetClosestPointTo(new Point3d(x, y, poly2d.Elevation), Vector3d.ZAxis, false);
                                                    double param1 = poly2d.GetParameterAtPoint(pt_on_line);
                                                    if (poly3d.EndParam < param1) param1 = poly3d.EndParam;
                                                    double sta1 = poly3d.GetDistanceAtParameter(param1);

                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = sta1;
                                                    }
                                                    else
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = sta1;
                                                    }

                                                    int index = populate_extra_columns_on_crossing_table(dt_crossing, dt_si, sta1);

                                                    if (lista_bands_for_generation.Contains(index) == false)
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1].Delete();
                                                    }
                                                }
                                                else
                                                {
                                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_2DSta] != DBNull.Value || _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)

                                                    {
                                                        dt_crossing.ImportRow(_AGEN_mainform.Data_Table_crossings.Rows[i]);

                                                        double sta1 = -123456;
                                                        if (_AGEN_mainform.Project_type == "2D")
                                                        {
                                                            sta1 = Convert.ToDouble(dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_2DSta]);

                                                        }
                                                        else
                                                        {
                                                            sta1 = Convert.ToDouble(dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_3DSta]);
                                                        }


                                                        if (_AGEN_mainform.COUNTRY == "CANADA")
                                                        {
                                                            if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                                            {
                                                                for (int j = 0; j < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++j)
                                                                {
                                                                    if (_AGEN_mainform.dt_centerline.Rows[j]["3DSta"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["3DSta"]).Replace("+", "")) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"]).Replace("+", "")) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j]["X"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["X"])) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j]["Y"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["Y"])) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j]["Z"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["Z"])) == true &&
                                                                         _AGEN_mainform.dt_centerline.Rows[j + 1]["X"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"])) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j + 1]["Y"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"])) == true &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j + 1]["Z"] != DBNull.Value &&
                                                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["Z"])) == true)
                                                                    {
                                                                        double sta_cl1 = Convert.ToDouble(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["3DSta"]).Replace("+", ""));
                                                                        double sta_cl2 = Convert.ToDouble(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                                        if (sta1 >= sta_cl1 && sta1 <= sta_cl2)
                                                                        {


                                                                            double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["X"]);
                                                                            double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Y"]);
                                                                            double z1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Z"]);
                                                                            double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"]);
                                                                            double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"]);
                                                                            double z2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Z"]);

                                                                            double x = x1 + (x2 - x1) * (sta1 - sta_cl1) / (sta_cl2 - sta_cl1);
                                                                            double y = y1 + (y2 - y1) * (sta1 - sta_cl1) / (sta_cl2 - sta_cl1);
                                                                            double z = z1 + (z2 - z1) * (sta1 - sta_cl1) / (sta_cl2 - sta_cl1);


                                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_x] = x;
                                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_y] = y;
                                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Elev] = z;

                                                                            j = _AGEN_mainform.dt_centerline.Rows.Count;

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {

                                                            if (sta1 < 0) sta1 = 0;
                                                            if (sta1 > poly3d.Length) sta1 = poly3d.Length - 0.000001;
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_x] = poly3d.GetPointAtDist(sta1).X;
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_y] = poly3d.GetPointAtDist(sta1).Y;
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Elev] = poly3d.GetPointAtDist(sta1).Z;
                                                        }


                                                        int index = populate_extra_columns_on_crossing_table(dt_crossing, dt_si, sta1);

                                                        if (lista_bands_for_generation.Contains(index) == false)
                                                        {
                                                            dt_crossing.Rows[dt_crossing.Rows.Count - 1].Delete();
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }

                                #region pi from centerline

                                string continut_tb_round = textBox_rounding_decimal_degrees.Text;

                                if (_AGEN_mainform.dt_centerline != null)
                                {
                                    if (_AGEN_mainform.dt_centerline.Rows.Count > 2)
                                    {
                                        for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++i)
                                        {

                                            if (_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAng] != DBNull.Value &&
                                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAngDMS] != DBNull.Value &&
                                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value &&
                                                _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                            {
                                                double defl = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAng]);
                                                string deflDMS = Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_DeflAngDMS]);

                                                string side1 = " RT";
                                                if (deflDMS.Contains("LT") == true)
                                                {
                                                    side1 = " LT";
                                                }

                                                    if (Functions.IsNumeric(continut_tb_round))
                                                {
                                                    double rr = Math.Abs(Convert.ToDouble(continut_tb_round));
                                                    double nr = Math.Round(defl / rr, 0);
                                                    defl = nr * rr;

                                                    deflDMS = Functions.Get_DMS(defl, 0)+ side1;
                                                }

                                                if (defl >= min_ang)
                                                {
                                                    dt_crossing.Rows.Add();
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Type] = _AGEN_mainform.crossing_type_pi;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_DeflAng] = defl;



                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.col_desc] = deflDMS;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_2DSta];
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta];
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_x] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_x];
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_y] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_y];
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Elev] = _AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_z];

                                                    double sta1 = -123456.7;
                                                    if (_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_2DSta] != DBNull.Value)
                                                    {
                                                        sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_2DSta]);
                                                    }

                                                    if (_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                                    {
                                                        sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i][_AGEN_mainform.Col_3DSta]);
                                                    }

                                                    int index = populate_extra_columns_on_crossing_table(dt_crossing, dt_si, sta1);

                                                    if (lista_bands_for_generation.Contains(index) == false)
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1].Delete();
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                                #endregion

                                #region property lines
                                if (checkBox_include_property_lines.Checked == true && _AGEN_mainform.COUNTRY == "USA")
                                {
                                    if (_AGEN_mainform.Data_Table_property != null)
                                    {
                                        if (_AGEN_mainform.Data_Table_property.Rows.Count > 1)
                                        {
                                            for (int i = 1; i < _AGEN_mainform.Data_Table_property.Rows.Count; ++i)
                                            {

                                                if (_AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"] != DBNull.Value
                                                     && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"])) == true
                                                     && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"])) == true)
                                                {
                                                    dt_crossing.Rows.Add();
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][col_xing_block] = block_prop_name;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][col_xing_sta] = atr_prop_sta;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][col_xing_descr] = atr_prop_descr;

                                                    double x = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["X_Beg"]);
                                                    double y = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["Y_Beg"]);

                                                    Point3d pt_on_line = poly2d.GetClosestPointTo(new Point3d(x, y, poly2d.Elevation), Vector3d.ZAxis, false);
                                                    double param1 = poly2d.GetParameterAtPoint(pt_on_line);
                                                    if (poly3d.EndParam < param1) param1 = poly3d.EndParam;

                                                    double sta1 = poly3d.GetDistanceAtParameter(param1);

                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = sta1;

                                                    }
                                                    else
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = sta1;
                                                    }

                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_x] = x;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_y] = y;
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.Col_Type] = "ownership";
                                                    dt_crossing.Rows[dt_crossing.Rows.Count - 1][_AGEN_mainform.col_desc] = "PROPERTY LINE";

                                                    int index = populate_extra_columns_on_crossing_table(dt_crossing, dt_si, sta1);

                                                    if (lista_bands_for_generation.Contains(index) == false)
                                                    {
                                                        dt_crossing.Rows[dt_crossing.Rows.Count - 1].Delete();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion

                                if (dt_crossing.Rows.Count == 0)
                                {
                                    _AGEN_mainform.tpage_processing.Hide();
                                    set_enable_true();
                                    MessageBox.Show("no data found");
                                    return;
                                }

                                if (_AGEN_mainform.Project_type == "2D")
                                {
                                    dt_crossing = Functions.Sort_data_table(dt_crossing, _AGEN_mainform.Col_2DSta);
                                }
                                else
                                {
                                    dt_crossing = Functions.Sort_data_table(dt_crossing, _AGEN_mainform.Col_3DSta);
                                }

                                #endregion

                                #region ADD MEASURED TO ST_EQ 
                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.COUNTRY == "USA")
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {
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
                                                double param1 = poly2d.GetParameterAtPoint(pt_on_2d);
                                                if (poly3d.EndParam < param1) param1 = poly3d.EndParam;

                                                double eq_meas = poly3d.GetDistanceAtParameter(param1);
                                                _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                                            }
                                        }
                                    }


                                }
                                #endregion

                                Functions.Creaza_layer(_AGEN_mainform.layer_crossing_band_text, 7, true);
                                Functions.Creaza_layer(_AGEN_mainform.layer_crossing_band_pi, 2, true);
                                Functions.Creaza_layer(lnp, 40, false);

                                if (checkBox_use_blocks.Checked == true)
                                {
                                    Functions.Creaza_layer(_AGEN_mainform.layer_crossing_band_matchline, 2, true);

                                }

                                string pi_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_pi_prefix();
                                string station_prefix = "";



                                List<Point3d> lista_puncte_corner = new List<Point3d>();
                                List<double> lista_ml_len = new List<double>();
                                List<double> lista_y = new List<double>();

                                #region draw crossings

                                #region OD DATA TABLE
                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                                Functions.Create_crossing_od_table();
                                #endregion

                                //Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_crossing);

                                #region draw rectangles

                                for (int i = 0; i < lista_bands_for_generation.Count; ++i)
                                {

                                    double M1_meas = Convert.ToDouble(dt_si.Rows[lista_bands_for_generation[i]]["meas1"]);
                                    double M2_meas = Convert.ToDouble(dt_si.Rows[lista_bands_for_generation[i]]["meas2"]);

                                    double M1_display = Convert.ToDouble(dt_si.Rows[lista_bands_for_generation[i]][_AGEN_mainform.Col_M1]);
                                    double M2_display = Convert.ToDouble(dt_si.Rows[lista_bands_for_generation[i]][_AGEN_mainform.Col_M2]);

                                    if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                    {
                                        M1_display = Functions.Station_equation_ofV2(M1_meas, _AGEN_mainform.dt_station_equation);
                                        M2_display = Functions.Station_equation_ofV2(M2_meas, _AGEN_mainform.dt_station_equation);
                                    }

                                    if (M2_meas >= poly3d.Length) M2_meas = poly3d.Length - 0.001;
                                    if (M2_meas < 0) M2_meas = 0;

                                    if (M1_meas >= poly3d.Length) M1_meas = poly3d.Length - 0.001;
                                    if (M1_meas < 0) M1_meas = 0;

                                    Point3d pm1 = poly3d.GetPointAtDist(M1_meas);
                                    Point3d pm2 = poly3d.GetPointAtDist(M2_meas);

                                    double x1 = pm1.X;
                                    double y1 = pm1.Y;
                                    double x2 = pm2.X;
                                    double y2 = pm2.Y;

                                    double ml_len = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);


                                    Polyline vp_vw1 = new Polyline();

                                    vp_vw1.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_cross.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw1.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_cross.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw1.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_cross.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw1.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_cross.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                    vp_vw1.Closed = true;
                                    vp_vw1.Layer = lnp;
                                    vp_vw1.ColorIndex = 3;//GREEN
                                    BTrecord.AppendEntity(vp_vw1);
                                    Trans1.AddNewlyCreatedDBObject(vp_vw1, true);

                                    Polyline vp_vw2 = new Polyline();

                                    vp_vw2.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_cross.X - ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw2.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_cross.X + ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw2.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_cross.X + ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                    vp_vw2.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_cross.X - ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                    vp_vw2.Closed = true;
                                    vp_vw2.Layer = lnp;
                                    vp_vw2.ColorIndex = 1;//RED
                                    BTrecord.AppendEntity(vp_vw2);
                                    Trans1.AddNewlyCreatedDBObject(vp_vw2, true);

                                    MText Band_label = new MText();
                                    Band_label.Contents = Convert.ToString(dt_si.Rows[lista_bands_for_generation[i]]["DwgNo"]);
                                    double textH = 1;
                                    if (Functions.Round_Down(_AGEN_mainform.Vw_cross_height / 3, 1) > 0) textH = Functions.Round_Down(_AGEN_mainform.Vw_cross_height / 3, 1);
                                    Band_label.TextHeight = textH;
                                    Band_label.Rotation = 0;
                                    Band_label.Attachment = AttachmentPoint.MiddleLeft;
                                    Band_label.Location = new Point3d(_AGEN_mainform.Point0_cross.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height / 2 - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation, 0);
                                    Band_label.Layer = lnp;
                                    Band_label.TextStyleId = Txt_style_id_standard;

                                    double gap1 = (_AGEN_mainform.Vw_width - ml_len * _AGEN_mainform.Vw_scale) / 2;
                                    Extents3d gerect = Band_label.GeometricExtents;
                                    Point3d p2 = gerect.MaxPoint;
                                    Point3d p1 = gerect.MinPoint;
                                    bool repeat1 = false;
                                    do
                                    {
                                        if (p2.X - p1.X > gap1 - TextH_pi_crossing && Band_label.TextHeight >= 2)
                                        {
                                            Band_label.TextHeight = Band_label.TextHeight - 1;
                                            repeat1 = true;
                                            gerect = Band_label.GeometricExtents;
                                            p2 = gerect.MaxPoint;
                                            p1 = gerect.MinPoint;
                                        }
                                        else
                                        {
                                            repeat1 = false;
                                        }
                                    }
                                    while (repeat1 == true);

                                    BTrecord.AppendEntity(Band_label);
                                    Trans1.AddNewlyCreatedDBObject(Band_label, true);

                                    if (_AGEN_mainform.XingDeltay1 == -123.456) _AGEN_mainform.XingDeltay1 = 2 * TextH_crossing;

                                    if (checkBox_use_blocks.Checked == true)
                                    {
                                        _AGEN_mainform.XingDeltay1 = 0;
                                    }

                                    double y0 = _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation - _AGEN_mainform.Vw_cross_height;
                                    double Y_m1 = y0 + _AGEN_mainform.XingDeltay1;
                                    double X_m1 = _AGEN_mainform.Point0_cross.X - lr * TextH_crossing - lr * ml_len * _AGEN_mainform.Vw_scale / 2;
                                    double X_m2 = _AGEN_mainform.Point0_cross.X + lr * TextH_crossing + lr * ml_len * _AGEN_mainform.Vw_scale / 2;

                                    string matchline1_chainage = Functions.Get_chainage_from_double(M1_display, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                    string matchline2_chainage = Functions.Get_chainage_from_double(M2_display, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                    string layer_label = lnp;

                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                    {
                                        layer_label = _AGEN_mainform.layer_crossing_band_text;
                                    }



                                    #region OD DATA TABLE
                                    List<object> Lista_val1 = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type1 = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                    if (segment1 == "not defined") segment1 = "";
                                    Lista_val1.Add(segment1);
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val1.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute + " by " + Environment.UserName);
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val1.Add(x1.ToString());
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val1.Add(y2.ToString());
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val1.Add(matchline1_chainage);
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val1.Add("matchline");
                                    Lista_type1.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    List<object> Lista_val2 = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type2 = new List<Autodesk.Gis.Map.Constants.DataType>();

                                    Lista_val2.Add(segment1);
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val2.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute + " by " + Environment.UserName);
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val2.Add(x2.ToString());
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val2.Add(y2.ToString());
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val2.Add(matchline2_chainage);
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val2.Add("matchline");
                                    Lista_type2.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    #endregion


                                    if (checkBox_use_blocks.Checked == false)
                                    {



                                        MText Crossing_M1 = new MText();
                                        Crossing_M1.Location = new Point3d(X_m1, Y_m1, 0);
                                        Crossing_M1.Layer = layer_label;
                                        Crossing_M1.TextStyleId = Txt_style_id_crossing;
                                        Crossing_M1.TextHeight = TextH_crossing;
                                        Crossing_M1.Rotation = Math.PI / 2;
                                        Crossing_M1.Contents = matchline1_chainage;
                                        Crossing_M1.Attachment = AttachmentPoint.MiddleLeft;
                                        BTrecord.AppendEntity(Crossing_M1);
                                        Trans1.AddNewlyCreatedDBObject(Crossing_M1, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_M1.ObjectId, "Agen_crossingV2", Lista_val1, Lista_type1);



                                        MText Crossing_M2 = new MText();
                                        Crossing_M2.Location = new Point3d(X_m2, Y_m1, 0);
                                        Crossing_M2.Layer = layer_label;
                                        Crossing_M2.TextStyleId = Txt_style_id_crossing;
                                        Crossing_M2.TextHeight = TextH_crossing;
                                        Crossing_M2.Rotation = Math.PI / 2;
                                        Crossing_M2.Contents = matchline2_chainage;
                                        Crossing_M2.Attachment = AttachmentPoint.MiddleLeft;
                                        BTrecord.AppendEntity(Crossing_M2);
                                        Trans1.AddNewlyCreatedDBObject(Crossing_M2, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_M2.ObjectId, "Agen_crossingV2", Lista_val2, Lista_type2);

                                        Extents3d gem2 = Crossing_M2.GeometricExtents;
                                        Point3d ptm1 = gem2.MaxPoint;
                                        Point3d ptm2 = gem2.MinPoint;

                                        double dy = Math.Abs(ptm1.Y - ptm2.Y);


                                        double scale1 = 1;
                                        if (_AGEN_mainform.XingDeltay2 == -123.456)
                                        {
                                            _AGEN_mainform.XingDeltay2 = _AGEN_mainform.XingDeltay1 + (dy + 2 * TextH_pi_crossing) * scale1;
                                        }

                                        MText matchline_M1 = new MText();
                                        matchline_M1.Location = new Point3d(X_m1, y0 + _AGEN_mainform.XingDeltay2, 0);
                                        matchline_M1.Layer = layer_label;
                                        matchline_M1.TextStyleId = Txt_style_id_crossing;
                                        matchline_M1.TextHeight = TextH_crossing;
                                        matchline_M1.Rotation = Math.PI / 2;
                                        matchline_M1.Contents = "MATCH LINE";
                                        matchline_M1.Attachment = AttachmentPoint.MiddleLeft;
                                        BTrecord.AppendEntity(matchline_M1);
                                        Trans1.AddNewlyCreatedDBObject(matchline_M1, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, matchline_M1.ObjectId, "Agen_crossingV2", Lista_val1, Lista_type1);



                                        MText matchline_M2 = new MText();
                                        matchline_M2.Location = new Point3d(X_m2, y0 + _AGEN_mainform.XingDeltay2, 0);
                                        matchline_M2.Layer = layer_label;
                                        matchline_M2.TextStyleId = Txt_style_id_crossing;
                                        matchline_M2.TextHeight = TextH_crossing;
                                        matchline_M2.Rotation = Math.PI / 2;
                                        matchline_M2.Contents = "MATCH LINE";
                                        matchline_M2.Attachment = AttachmentPoint.MiddleLeft;
                                        BTrecord.AppendEntity(matchline_M2);
                                        Trans1.AddNewlyCreatedDBObject(matchline_M2, true);
                                        Functions.Populate_object_data_table_from_objectid(Tables1, matchline_M2.ObjectId, "Agen_crossingV2", Lista_val2, Lista_type2);


                                        Point2d p2d1 = new Point2d(_AGEN_mainform.Point0_cross.X - ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation);
                                        Point2d p2d2 = new Point2d(_AGEN_mainform.Point0_cross.X - ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation);
                                        Point2d p2d3 = new Point2d(_AGEN_mainform.Point0_cross.X + ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation);
                                        Point2d p2d4 = new Point2d(_AGEN_mainform.Point0_cross.X + ml_len * _AGEN_mainform.Vw_scale / 2, _AGEN_mainform.Point0_cross.Y - _AGEN_mainform.Vw_cross_height - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation);

                                        Polyline poly_m1 = new Polyline();
                                        poly_m1.AddVertexAt(0, p2d1, 0, 0, 0);
                                        poly_m1.AddVertexAt(1, p2d2, 0, 0, 0);
                                        poly_m1.Layer = layer_label;
                                        poly_m1.ColorIndex = 256;
                                        BTrecord.AppendEntity(poly_m1);
                                        Trans1.AddNewlyCreatedDBObject(poly_m1, true);

                                        if (lr == 1)
                                        {
                                            Functions.Populate_object_data_table_from_objectid(Tables1, poly_m1.ObjectId, "Agen_crossingV2", Lista_val1, Lista_type1);
                                        }
                                        else
                                        {
                                            Functions.Populate_object_data_table_from_objectid(Tables1, poly_m1.ObjectId, "Agen_crossingV2", Lista_val2, Lista_type2);
                                        }



                                        Polyline poly_m2 = new Polyline();
                                        poly_m2.AddVertexAt(0, p2d3, 0, 0, 0);
                                        poly_m2.AddVertexAt(1, p2d4, 0, 0, 0);
                                        poly_m2.Layer = layer_label;
                                        poly_m2.ColorIndex = 256;
                                        BTrecord.AppendEntity(poly_m2);
                                        Trans1.AddNewlyCreatedDBObject(poly_m2, true);

                                        if (lr == 1)
                                        {
                                            Functions.Populate_object_data_table_from_objectid(Tables1, poly_m2.ObjectId, "Agen_crossingV2", Lista_val2, Lista_type2);
                                        }
                                        else
                                        {
                                            Functions.Populate_object_data_table_from_objectid(Tables1, poly_m2.ObjectId, "Agen_crossingV2", Lista_val1, Lista_type1);
                                        }
                                    }


                                    if (checkBox_use_blocks.Checked == true)
                                    {
                                        string block_name = "";
                                        string atr_sta = "";
                                        string atr_descr = "";

                                        string visib1 = "";

                                       


                                        if (dt_crossing != null && dt_crossing.Rows.Count > 0)
                                        {
                                            for (int k = 0; k < dt_crossing.Rows.Count; ++k)
                                            {
                                                if (atr_sta == "")
                                                {
                                                    if (dt_crossing.Rows[k][col_xing_sta] != DBNull.Value)
                                                    {
                                                        atr_sta = Convert.ToString(dt_crossing.Rows[k][col_xing_sta]);
                                                    }
                                                }
                                                if (atr_descr == "")
                                                {
                                                    if (dt_crossing.Rows[k][col_xing_descr] != DBNull.Value)
                                                    {
                                                        atr_descr = Convert.ToString(dt_crossing.Rows[k][col_xing_descr]);
                                                    }
                                                }
                                                if (block_name == "")
                                                {
                                                    if (dt_crossing.Rows[k][col_xing_block] != DBNull.Value)
                                                    {
                                                        block_name = Convert.ToString(dt_crossing.Rows[k][col_xing_block]);
                                                    }
                                                }

                                                if (dt_crossing.Rows[k][col_visibility] != DBNull.Value)
                                                {
                                                    visib1 = Convert.ToString(dt_crossing.Rows[k][col_visibility]);
                                                }


                                                if (block_name != "" && atr_sta != "") k = dt_crossing.Rows.Count;

                                            }

                                            if (block_name != "")
                                            {
                                                System.Collections.Specialized.StringCollection col_num1 = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val1 = new System.Collections.Specialized.StringCollection();

                                                col_num1.Add(atr_sta);
                                                col_val1.Add(matchline1_chainage);
                                                col_num1.Add(atr_descr);
                                                col_val1.Add("MATCH LINE");

                                                col_num1.Add(atr_sta + "1");
                                                col_val1.Add(matchline1_chainage);
                                                col_num1.Add(atr_descr + "1");
                                                col_val1.Add("MATCH LINE");

                                                col_num1.Add(atr_sta + "11");
                                                col_val1.Add(matchline1_chainage);
                                                col_num1.Add(atr_descr + "11");
                                                col_val1.Add("MATCH LINE");

                                                col_num1.Add(atr_sta + "111");
                                                col_val1.Add(matchline1_chainage);
                                                col_num1.Add(atr_descr + "111");
                                                col_val1.Add("MATCH LINE");

                                                col_num1.Add(atr_sta + "1111");
                                                col_val1.Add(matchline1_chainage);
                                                col_num1.Add(atr_descr + "1111");
                                                col_val1.Add("MATCH LINE");

                                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                                                                            block_name, new Point3d(X_m1, y0, 0), 1, 0, _AGEN_mainform.layer_crossing_band_matchline, col_num1, col_val1);
                                                Functions.set_block_visibility(Block1, visib1);

                                                System.Collections.Specialized.StringCollection col_num2 = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection col_val2 = new System.Collections.Specialized.StringCollection();

                                                col_num2.Add(atr_sta);
                                                col_val2.Add(matchline2_chainage);
                                                col_num2.Add(atr_descr);
                                                col_val2.Add("MATCH LINE");

                                                col_num2.Add(atr_sta + "1");
                                                col_val2.Add(matchline2_chainage);
                                                col_num2.Add(atr_descr + "1");
                                                col_val2.Add("MATCH LINE");

                                                col_num2.Add(atr_sta + "11");
                                                col_val2.Add(matchline2_chainage);
                                                col_num2.Add(atr_descr + "11");
                                                col_val2.Add("MATCH LINE");

                                                col_num2.Add(atr_sta + "111");
                                                col_val2.Add(matchline2_chainage);
                                                col_num2.Add(atr_descr + "111");
                                                col_val2.Add("MATCH LINE");

                                                col_num2.Add(atr_sta + "1111");
                                                col_val2.Add(matchline2_chainage);
                                                col_num2.Add(atr_descr + "1111");
                                                col_val2.Add("MATCH LINE");

                                                BlockReference Block2 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                                                                            block_name, new Point3d(X_m2, y0, 0), 1, 0, _AGEN_mainform.layer_crossing_band_matchline, col_num2, col_val2);
                                                Functions.set_block_visibility(Block2, visib1);

                                            }

                                        }

                                    }



                                    if (_AGEN_mainform.XingDeltay1 == -123.456) _AGEN_mainform.XingDeltay1 = 2 * TextH_crossing;

                                    lista_puncte_corner.Add(new Point3d(_AGEN_mainform.Point0_cross.X - lr * ml_len * _AGEN_mainform.Vw_scale / 2,
                                                       _AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation - _AGEN_mainform.Vw_cross_height + _AGEN_mainform.XingDeltay1,
                                                       0));
                                    lista_ml_len.Add(ml_len);

                                    lista_y.Add(_AGEN_mainform.Point0_cross.Y - lista_bands_for_generation[i] * _AGEN_mainform.Band_Separation - _AGEN_mainform.Vw_cross_height);

                                }
                                #endregion



                                Point3d Prevpt = new Point3d();
                                Point3d PtM1 = new Point3d();
                                int rand_prev = -1;

                                // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_crossing);

                                for (int i = 0; i < dt_crossing.Rows.Count; ++i)
                                {
                                    debug_i = i;
                                    double y0 = -123.456;
                                    int nr_rand = Convert.ToInt32(dt_crossing.Rows[i]["index1"]);
                                    y0 = lista_y[lista_bands_for_generation.IndexOf(nr_rand)];

                                    if (nr_rand != rand_prev)
                                    {
                                        PtM1 = lista_puncte_corner[lista_bands_for_generation.IndexOf(nr_rand)];
                                        Prevpt = PtM1;

                                    }

                                    double M1 = Convert.ToDouble(dt_si.Rows[nr_rand]["meas1"]);
                                    double M2 = Convert.ToDouble(dt_si.Rows[nr_rand]["meas2"]);

                                    double ml_len = lista_ml_len[lista_bands_for_generation.IndexOf(nr_rand)];
                                    string dwg_name = Convert.ToString(dt_crossing.Rows[i]["dwg"]);
                                    double rectangle_corner_x = _AGEN_mainform.Point0_cross.X - lr * ml_len * _AGEN_mainform.Vw_scale / 2;

                                    double Sta3d = -1;

                                    if (_AGEN_mainform.Project_type == "2D")
                                    {
                                        Sta3d = Convert.ToDouble(dt_crossing.Rows[i][_AGEN_mainform.Col_2DSta]);
                                    }
                                    else
                                    {
                                        Sta3d = Convert.ToDouble(dt_crossing.Rows[i][_AGEN_mainform.Col_3DSta]);
                                    }

                                    Sta3d = Math.Round(Sta3d, _AGEN_mainform.round1);
                                    if (Sta3d >= poly3d.Length) Sta3d = poly3d.Length - 0.0001;


                                    string Type1 = Convert.ToString(dt_crossing.Rows[i][_AGEN_mainform.Col_Type]);
                                    string Desc1 = Convert.ToString(dt_crossing.Rows[i][_AGEN_mainform.col_desc]);

                                    double X = Convert.ToDouble(dt_crossing.Rows[i]["X"]);
                                    double Y = Convert.ToDouble(dt_crossing.Rows[i]["Y"]);
                                    Point3d pt1 = poly2d.GetClosestPointTo(new Point3d(X, Y, poly2d.Elevation), Vector3d.ZAxis, false);
                                    double par1 = poly2d.GetParameterAtPoint(pt1);
                                    if (poly3d.EndParam < par1) par1 = poly2d.EndParam;

                                    double dist2d = poly2d.GetDistanceAtParameter(par1);

                                    #region OD DATA TABLE
                                    List<object> Lista_val = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                    if (segment1 == "not defined") segment1 = "";
                                    Lista_val.Add(segment1);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute + " by " + Environment.UserName);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Lista_val.Add(X.ToString());
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Lista_val.Add(Y.ToString());
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                                    #endregion

                                    if (poly3d.Length <= M2)
                                    {
                                        M2 = poly3d.Length - 0.0001;
                                    }

                                    Point3d pm1 = poly3d.GetPointAtDist(M1);
                                    Point3d pm2 = poly3d.GetPointAtDist(M2);
                                    double xm1 = pm1.X;
                                    double ym1 = pm1.Y;
                                    double xm2 = pm2.X;
                                    double ym2 = pm2.Y;

                                    if (Type1 == _AGEN_mainform.crossing_type_pi || Type1.ToLower() == "xpi")
                                    {
                                        if (dt_crossing.Rows[i][_AGEN_mainform.Col_DeflAng] != DBNull.Value)
                                        {
                                            double defl1 = Convert.ToDouble(dt_crossing.Rows[i][_AGEN_mainform.Col_DeflAng]);
                                            if (defl1 >= min_ang)
                                            {
                                                Point3d P1 = poly3d.GetPointAtDist(Sta3d);
                                                Line LineM1M2 = new Line(pm1, pm2);
                                                Point3d PP1 = LineM1M2.GetClosestPointTo(P1, Vector3d.ZAxis, false);

                                                double Deltax = pm1.DistanceTo(PP1) * _AGEN_mainform.Vw_scale;
                                                Point3d Inspt = new Point3d(PtM1.X + lr * Deltax, Prevpt.Y, 0);

                                                if (lr == 1)
                                                {
                                                    if (Inspt.X < Prevpt.X)
                                                    {
                                                        Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                    }
                                                    else
                                                    {
                                                        if (Inspt.X - Prevpt.X < min_dist)
                                                        {
                                                            Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (Inspt.X > Prevpt.X)
                                                    {
                                                        Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                    }
                                                    else
                                                    {
                                                        if (-Inspt.X + Prevpt.X < min_dist)
                                                        {
                                                            Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                        }
                                                    }
                                                }

                                                string sta_string = "";
                                                if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_display_station_value() == true)
                                                {
                                                    if (_AGEN_mainform.COUNTRY == "USA")
                                                    {
                                                        sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";
                                                        if (checkBox_use_blocks.Checked == true)
                                                        {
                                                            sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                                        }
                                                    }
                                                    if (_AGEN_mainform.COUNTRY == "CANADA")
                                                    {
                                                        double b1 = -1.23456;
                                                        double display_sta = Functions.get_stationCSF_from_point(poly2d, pt1, dist2d, _AGEN_mainform.dt_centerline, ref b1);

                                                        sta_string = Functions.Get_chainage_from_double(display_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";
                                                        if (checkBox_use_blocks.Checked == true)
                                                        {
                                                            sta_string = Functions.Get_chainage_from_double(display_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                        }
                                                    }
                                                    station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();
                                                }

                                                string a = "";
                                                string c = "";
                                                if (Wfactor == 1)
                                                {
                                                    if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                    {
                                                        a = "\\L{";
                                                        c = "}";
                                                    }
                                                }
                                                else
                                                {
                                                    if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                    {
                                                        a = "{\\W" + Wfactor.ToString() + ";\\L";
                                                        c = "}";
                                                    }
                                                    else
                                                    {
                                                        a = "{\\W" + Wfactor.ToString() + ";";
                                                        c = "}";
                                                    }
                                                }

                                                station_prefix = Functions.remove_space_from_start_and_end_of_a_string(station_prefix);
                                                sta_string = Functions.remove_space_from_start_and_end_of_a_string(sta_string);
                                                pi_prefix = Functions.remove_space_from_start_and_end_of_a_string(pi_prefix);
                                                Desc1 = Functions.remove_space_from_start_and_end_of_a_string(Desc1);

                                                if (station_prefix.Length > 0)
                                                {
                                                    station_prefix = station_prefix + " ";
                                                }

                                                if (pi_prefix.Length > 0)
                                                {
                                                    pi_prefix = pi_prefix + " ";
                                                }

                                                string Continut1 = a + station_prefix + sta_string + " " + pi_prefix + Desc1 + c;
                                                Continut1 = Functions.remove_space_from_start_and_end_of_a_string(Continut1);

                                                #region OD Table
                                                Lista_val.Add(sta_string);
                                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                                Lista_val.Add(Desc1);
                                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                                #endregion

                                                ObjectId obid = ObjectId.Null;
                                                if (checkBox_use_blocks.Checked == false)
                                                {
                                                    if (checkBox_split_station.Checked == false)
                                                    {
                                                        MText Crossing1 = new MText();
                                                        Crossing1.Location = Inspt;
                                                        Crossing1.Layer = _AGEN_mainform.layer_crossing_band_pi;
                                                        Crossing1.TextStyleId = Txt_style_id_pi_crossing;
                                                        Crossing1.TextHeight = TextH_pi_crossing;
                                                        Crossing1.Rotation = TextR;
                                                        Crossing1.Contents = Continut1;
                                                        BTrecord.AppendEntity(Crossing1);
                                                        Trans1.AddNewlyCreatedDBObject(Crossing1, true);
                                                        Functions.Populate_object_data_table_from_objectid(Tables1, Crossing1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                        obid = Crossing1.ObjectId;
                                                    }

                                                    #region SPLIT STATION FROM DESCRIPTION
                                                    if (checkBox_split_station.Checked == true)
                                                    {
                                                        string Continut2 = a + station_prefix + " " + sta_string + c;
                                                        Continut2 = Functions.remove_space_from_start_and_end_of_a_string(Continut2);

                                                        string Continut3 = a + pi_prefix + " " + Desc1 + c;
                                                        Continut3 = Functions.remove_space_from_start_and_end_of_a_string(Continut3);

                                                        MText Crossing_sta1 = new MText();
                                                        Crossing_sta1.Location = Inspt;
                                                        Crossing_sta1.Layer = _AGEN_mainform.layer_crossing_band_pi;
                                                        Crossing_sta1.TextStyleId = Txt_style_id_pi_crossing;
                                                        Crossing_sta1.TextHeight = TextH_pi_crossing;
                                                        Crossing_sta1.Rotation = TextR;
                                                        Crossing_sta1.Contents = Continut2;
                                                        Crossing_sta1.Attachment = AttachmentPoint.MiddleLeft;
                                                        BTrecord.AppendEntity(Crossing_sta1);
                                                        Trans1.AddNewlyCreatedDBObject(Crossing_sta1, true);
                                                        Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_sta1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                        obid = Crossing_sta1.ObjectId;

                                                        Extents3d gerect = Crossing_sta1.GeometricExtents;
                                                        Point3d p1 = gerect.MaxPoint;
                                                        Point3d p2 = gerect.MinPoint;

                                                        double dy = Math.Abs(p1.Y - p2.Y);
                                                        double scale1 = 1;

                                                        if (_AGEN_mainform.XingDeltay2 == -123.456)
                                                        {
                                                            _AGEN_mainform.XingDeltay2 = _AGEN_mainform.XingDeltay1 + (dy + 2 * TextH_pi_crossing) * scale1;
                                                        }

                                                        if (checkBox_draw_angle_symbol.Checked == true)
                                                        {
                                                            Polyline Poly_arc = new Polyline();
                                                            Poly_arc.AddVertexAt(0, new Point2d(Inspt.X - 3.377 * scale1 + TextH_crossing / 2, y0 + _AGEN_mainform.XingDeltay2 + (2.446) * scale1), -Math.Tan((117 * Math.PI / 180) / 4), 0, 0);
                                                            Poly_arc.AddVertexAt(1, new Point2d(Inspt.X + 0.997 * scale1 + TextH_crossing / 2, y0 + _AGEN_mainform.XingDeltay2 + (3.092) * scale1), 0, 0, 0);
                                                            Poly_arc.Layer = _AGEN_mainform.layer_crossing_band_pi;
                                                            BTrecord.AppendEntity(Poly_arc);
                                                            Trans1.AddNewlyCreatedDBObject(Poly_arc, true);
                                                            Functions.Populate_object_data_table_from_objectid(Tables1, Poly_arc.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);

                                                            Polyline Poly_l = new Polyline();
                                                            Poly_l.AddVertexAt(0, new Point2d(Inspt.X + TextH_crossing / 2, y0 + _AGEN_mainform.XingDeltay2 + 4.16 * scale1), 0, 0, 0);
                                                            Poly_l.AddVertexAt(1, new Point2d(Inspt.X + TextH_crossing / 2, y0 + _AGEN_mainform.XingDeltay2), 0, 0, 0);
                                                            Poly_l.AddVertexAt(2, new Point2d(Inspt.X - 2.8 * scale1 + TextH_crossing / 2, y0 + _AGEN_mainform.XingDeltay2 + 4.16 * scale1), 0, 0, 0);
                                                            Poly_l.Layer = _AGEN_mainform.layer_crossing_band_pi;

                                                            BTrecord.AppendEntity(Poly_l);
                                                            Trans1.AddNewlyCreatedDBObject(Poly_l, true);
                                                            Functions.Populate_object_data_table_from_objectid(Tables1, Poly_l.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                            dy = dy + TextH_pi_crossing;
                                                        }

                                                        if (_AGEN_mainform.XingDeltay3 == -123.456)
                                                        {
                                                            _AGEN_mainform.XingDeltay3 = _AGEN_mainform.XingDeltay1 + dy + 4 * TextH_pi_crossing;
                                                        }

                                                        MText Crossing_desc1 = new MText();
                                                        Crossing_desc1.Location = new Point3d(Inspt.X, y0 + _AGEN_mainform.XingDeltay3, 0);
                                                        Crossing_desc1.Layer = _AGEN_mainform.layer_crossing_band_pi;
                                                        Crossing_desc1.TextStyleId = Txt_style_id_pi_crossing;
                                                        Crossing_desc1.TextHeight = TextH_pi_crossing;
                                                        Crossing_desc1.Rotation = TextR;
                                                        Crossing_desc1.Contents = Continut3;
                                                        Crossing_desc1.Attachment = AttachmentPoint.MiddleLeft;
                                                        BTrecord.AppendEntity(Crossing_desc1);
                                                        Trans1.AddNewlyCreatedDBObject(Crossing_desc1, true);
                                                        Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_desc1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                        obid = Crossing_desc1.ObjectId;
                                                    }
                                                    #endregion
                                                }

                                                if (checkBox_use_blocks.Checked == true)
                                                {
                                                    System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                                    System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                                    Colectie_nume_atribute.Add(atr_pi_sta);
                                                    Colectie_valori.Add(sta_string);
                                                    Colectie_nume_atribute.Add(atr_pi_descr);
                                                    Colectie_valori.Add(Desc1);

                                                    Colectie_nume_atribute.Add(atr_pi_sta + "1");
                                                    Colectie_valori.Add(sta_string);
                                                    Colectie_nume_atribute.Add(atr_pi_descr + "1");
                                                    Colectie_valori.Add(Desc1);

                                                    Colectie_nume_atribute.Add(atr_pi_sta + "11");
                                                    Colectie_valori.Add(sta_string);
                                                    Colectie_nume_atribute.Add(atr_pi_descr + "11");
                                                    Colectie_valori.Add(Desc1);


                                                    Colectie_nume_atribute.Add(atr_pi_sta + "111");
                                                    Colectie_valori.Add(sta_string);
                                                    Colectie_nume_atribute.Add(atr_pi_descr + "111");
                                                    Colectie_valori.Add(Desc1);


                                                    Colectie_nume_atribute.Add(atr_pi_sta + "1111");
                                                    Colectie_valori.Add(sta_string);
                                                    Colectie_nume_atribute.Add(atr_pi_descr + "1111");
                                                    Colectie_valori.Add(Desc1);

                                                    BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block_pi_name, Inspt, 1, 0, _AGEN_mainform.layer_crossing_band_text, Colectie_nume_atribute, Colectie_valori);

                                                }

                                                Prevpt = Inspt;
                                            }
                                        }
                                        else
                                        {
                                            Point3d P1 = poly3d.GetPointAtDist(Sta3d);
                                            Line LineM1M2 = new Line(pm1, pm2);
                                            Point3d PP1 = LineM1M2.GetClosestPointTo(P1, Vector3d.ZAxis, false);
                                            double Deltax = pm1.DistanceTo(PP1) * _AGEN_mainform.Vw_scale;
                                            Point3d Inspt = new Point3d(PtM1.X + lr * Deltax, Prevpt.Y, 0);

                                            if (lr == 1)
                                            {
                                                if (Inspt.X < Prevpt.X)
                                                {
                                                    Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                }
                                                else
                                                {
                                                    if (Inspt.X - Prevpt.X < min_dist)
                                                    {
                                                        Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (Inspt.X > Prevpt.X)
                                                {
                                                    Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                }
                                                else
                                                {
                                                    if (-Inspt.X + Prevpt.X < min_dist)
                                                    {
                                                        Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                    }
                                                }
                                            }

                                            string sta_string = "";
                                            string sta = "";
                                            if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_display_station_value() == true)
                                            {
                                                station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();
                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                {
                                                    sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";
                                                    sta = sta_string;
                                                }
                                                if (_AGEN_mainform.COUNTRY == "CANADA")
                                                {
                                                    double b1 = -1.23456;
                                                    double display_sta = Functions.get_stationCSF_from_point(poly2d, pt1, dist2d, _AGEN_mainform.dt_centerline, ref b1);
                                                    sta_string = Functions.Get_chainage_from_double(display_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";
                                                    sta = sta_string;
                                                }
                                            }

                                            string a = "";
                                            string c = "";
                                            if (Wfactor == 1)
                                            {
                                                if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                {
                                                    a = "\\L{";
                                                    c = "}";
                                                }
                                            }
                                            else
                                            {
                                                if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_pi_underline_value() == true)
                                                {
                                                    a = "{\\W" + Wfactor.ToString() + ";\\L";
                                                    c = "}";
                                                }
                                                else
                                                {
                                                    a = "{\\W" + Wfactor.ToString() + ";";
                                                    c = "}";
                                                }
                                            }

                                            sta_string = Functions.remove_space_from_start_and_end_of_a_string(sta_string);
                                            station_prefix = Functions.remove_space_from_start_and_end_of_a_string(station_prefix);
                                            Desc1 = Functions.remove_space_from_start_and_end_of_a_string(Desc1);

                                            if (station_prefix.Length > 0)
                                            {
                                                station_prefix = station_prefix + " ";
                                            }

                                            string Continut1 = a + station_prefix + sta_string + " " + Desc1 + c;
                                            Continut1 = Functions.remove_space_from_start_and_end_of_a_string(Continut1);

                                            #region OD Table
                                            Lista_val.Add(sta_string);
                                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                            Lista_val.Add(Desc1);
                                            Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                            #endregion

                                            ObjectId obid = ObjectId.Null;

                                            if (checkBox_use_blocks.Checked == false)
                                            {
                                                MText Crossing1 = new MText();
                                                Crossing1.Location = Inspt;
                                                Crossing1.Layer = _AGEN_mainform.layer_crossing_band_pi;
                                                Crossing1.TextStyleId = Txt_style_id_pi_crossing;
                                                Crossing1.TextHeight = TextH_crossing;
                                                Crossing1.Rotation = TextR;
                                                Crossing1.Contents = Continut1;
                                                BTrecord.AppendEntity(Crossing1);
                                                Trans1.AddNewlyCreatedDBObject(Crossing1, true);
                                                Functions.Populate_object_data_table_from_objectid(Tables1, Crossing1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                obid = Crossing1.ObjectId;
                                            }




                                            if (checkBox_use_blocks.Checked == true)
                                            {
                                                System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                                System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                                if (dt_crossing.Rows[i][col_xing_sta] != DBNull.Value)
                                                {
                                                    atr_sta = Convert.ToString(dt_crossing.Rows[i][col_xing_sta]);
                                                }
                                                if (dt_crossing.Rows[i][col_xing_descr] != DBNull.Value)
                                                {
                                                    atr_descr = Convert.ToString(dt_crossing.Rows[i][col_xing_descr]);
                                                }
                                                string block1 = block_xing_name;
                                                if (dt_crossing.Rows[i][col_xing_block] != DBNull.Value)
                                                {
                                                    block1 = Convert.ToString(dt_crossing.Rows[i][col_xing_block]);
                                                }
                                                Colectie_nume_atribute.Add(atr_sta);
                                                Colectie_valori.Add(sta);
                                                Colectie_nume_atribute.Add(atr_descr);
                                                Colectie_valori.Add(Desc1);

                                                Colectie_nume_atribute.Add(atr_sta + "1");
                                                Colectie_valori.Add(sta);
                                                Colectie_nume_atribute.Add(atr_descr + "1");
                                                Colectie_valori.Add(Desc1);

                                                Colectie_nume_atribute.Add(atr_sta + "11");
                                                Colectie_valori.Add(sta);
                                                Colectie_nume_atribute.Add(atr_descr + "11");
                                                Colectie_valori.Add(Desc1);

                                                Colectie_nume_atribute.Add(atr_sta + "111");
                                                Colectie_valori.Add(sta);
                                                Colectie_nume_atribute.Add(atr_descr + "111");
                                                Colectie_valori.Add(Desc1);

                                                Colectie_nume_atribute.Add(atr_sta + "1111");
                                                Colectie_valori.Add(sta);
                                                Colectie_nume_atribute.Add(atr_descr + "1111");
                                                Colectie_valori.Add(Desc1);


                                                BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block1, Inspt, 1, 0, _AGEN_mainform.layer_crossing_band_text, Colectie_nume_atribute, Colectie_valori);
                                                string visib1 = "";
                                                if (dt_crossing.Rows[i][col_visibility] != DBNull.Value)
                                                {
                                                    visib1 = Convert.ToString(dt_crossing.Rows[i][col_visibility]);
                                                }
                                                Functions.set_block_visibility(Block1, visib1);
                                            }
                                            Prevpt = Inspt;
                                        }
                                    }
                                    else
                                    {
                                        Point3d P1 = poly3d.GetPointAtDist(Sta3d);
                                        Line LineM1M2 = new Line(pm1, pm2);
                                        Point3d PP1 = LineM1M2.GetClosestPointTo(P1, Vector3d.ZAxis, false);
                                        double Deltax = pm1.DistanceTo(PP1) * _AGEN_mainform.Vw_scale;
                                        Point3d Inspt = new Point3d(PtM1.X + lr * Deltax, Prevpt.Y, 0);

                                        if (lr == 1)
                                        {
                                            if (Inspt.X < Prevpt.X)
                                            {
                                                Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                            }
                                            else
                                            {
                                                if (Inspt.X - Prevpt.X < min_dist)
                                                {
                                                    Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            if (Inspt.X > Prevpt.X)
                                            {
                                                Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                            }
                                            else
                                            {
                                                if (-Inspt.X + Prevpt.X < min_dist)
                                                {
                                                    Inspt = new Point3d(Prevpt.X + lr * min_dist, Prevpt.Y, 0);
                                                }
                                            }
                                        }

                                        string sta_string = "";
                                        string sta = "";
                                        if (_AGEN_mainform.tpage_crossing_draw.get_checkBox_display_station_value() == true)
                                        {
                                            station_prefix = _AGEN_mainform.tpage_crossing_draw.get_textBox_station_prefix();
                                            if (_AGEN_mainform.COUNTRY == "USA")
                                            {
                                                sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";
                                                if (checkBox_use_blocks.Checked == true)
                                                {
                                                    sta_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Sta3d, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                }
                                                sta = sta_string;
                                            }
                                            if (_AGEN_mainform.COUNTRY == "CANADA")
                                            {
                                                double b1 = -1.23456;
                                                double display_sta = Functions.get_stationCSF_from_point(poly2d, pt1, dist2d, _AGEN_mainform.dt_centerline, ref b1);
                                                sta_string = Functions.Get_chainage_from_double(display_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + " ";

                                                if (checkBox_use_blocks.Checked == true)
                                                {
                                                    sta_string = Functions.Get_chainage_from_double(display_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                }
                                                sta = sta_string;
                                            }
                                        }

                                        string a = "";
                                        string c = "";
                                        if (Wfactor == 1)
                                        {

                                        }
                                        else
                                        {
                                            a = "{\\W" + Wfactor.ToString() + ";";
                                            c = "}";
                                        }

                                        sta_string = Functions.remove_space_from_start_and_end_of_a_string(sta_string);
                                        station_prefix = Functions.remove_space_from_start_and_end_of_a_string(station_prefix);
                                        Desc1 = Functions.remove_space_from_start_and_end_of_a_string(Desc1);

                                        if (station_prefix.Length > 0)
                                        {
                                            station_prefix = station_prefix + " ";
                                        }

                                        string Continut1 = a + station_prefix + sta_string + " " + Desc1 + c;
                                        Continut1 = Functions.remove_space_from_start_and_end_of_a_string(Continut1);

                                        #region OD Table
                                        Lista_val.Add(sta_string);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                        Lista_val.Add(Desc1);
                                        Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                        #endregion

                                        ObjectId obid = ObjectId.Null;

                                        if (checkBox_use_blocks.Checked == false)
                                        {
                                            if (checkBox_split_station.Checked == false)
                                            {
                                                MText Crossing1 = new MText();
                                                Crossing1.Location = Inspt;
                                                Crossing1.Layer = _AGEN_mainform.layer_crossing_band_text;
                                                Crossing1.TextStyleId = Txt_style_id_crossing;
                                                Crossing1.TextHeight = TextH_crossing;
                                                Crossing1.Rotation = TextR;
                                                Crossing1.Contents = Continut1;
                                                BTrecord.AppendEntity(Crossing1);
                                                Trans1.AddNewlyCreatedDBObject(Crossing1, true);
                                                Functions.Populate_object_data_table_from_objectid(Tables1, Crossing1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                obid = Crossing1.ObjectId;
                                            }

                                            if (checkBox_split_station.Checked == true)
                                            {
                                                string Continut2 = a + station_prefix + " " + sta_string + c;
                                                Continut2 = Functions.remove_space_from_start_and_end_of_a_string(Continut2);

                                                string Continut3 = a + Desc1 + c;
                                                Continut3 = Functions.remove_space_from_start_and_end_of_a_string(Continut3);

                                                MText Crossing_sta1 = new MText();
                                                Crossing_sta1.Location = Inspt;
                                                Crossing_sta1.Layer = _AGEN_mainform.layer_crossing_band_text;
                                                Crossing_sta1.TextStyleId = Txt_style_id_crossing;
                                                Crossing_sta1.TextHeight = TextH_crossing;
                                                Crossing_sta1.Rotation = TextR;
                                                Crossing_sta1.Contents = Continut2;
                                                Crossing_sta1.Attachment = AttachmentPoint.MiddleLeft;
                                                BTrecord.AppendEntity(Crossing_sta1);
                                                Trans1.AddNewlyCreatedDBObject(Crossing_sta1, true);
                                                Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_sta1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                obid = Crossing_sta1.ObjectId;

                                                Extents3d gerect = Crossing_sta1.GeometricExtents;
                                                Point3d p1 = gerect.MaxPoint;
                                                Point3d p2 = gerect.MinPoint;

                                                double dy = Math.Abs(p1.Y - p2.Y);

                                                if (_AGEN_mainform.XingDeltay2 == -123.456)
                                                {
                                                    _AGEN_mainform.XingDeltay2 = _AGEN_mainform.XingDeltay1 + (dy + 2 * TextH_pi_crossing);
                                                }

                                                MText Crossing_desc1 = new MText();
                                                Crossing_desc1.Location = new Point3d(Inspt.X, y0 + _AGEN_mainform.XingDeltay2, 0);
                                                Crossing_desc1.Layer = _AGEN_mainform.layer_crossing_band_text;
                                                Crossing_desc1.TextStyleId = Txt_style_id_crossing;
                                                Crossing_desc1.TextHeight = TextH_crossing;
                                                Crossing_desc1.Rotation = TextR;
                                                Crossing_desc1.Contents = Continut3;
                                                Crossing_desc1.Attachment = AttachmentPoint.MiddleLeft;
                                                BTrecord.AppendEntity(Crossing_desc1);
                                                Trans1.AddNewlyCreatedDBObject(Crossing_desc1, true);
                                                Functions.Populate_object_data_table_from_objectid(Tables1, Crossing_desc1.ObjectId, "Agen_crossingV2", Lista_val, Lista_type);
                                                obid = Crossing_desc1.ObjectId;
                                            }
                                        }

                                        if (checkBox_use_blocks.Checked == true)
                                        {
                                            System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                            System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                            if (dt_crossing.Rows[i][col_xing_sta] != DBNull.Value)
                                            {
                                                atr_sta = Convert.ToString(dt_crossing.Rows[i][col_xing_sta]);
                                            }
                                            if (dt_crossing.Rows[i][col_xing_descr] != DBNull.Value)
                                            {
                                                atr_descr = Convert.ToString(dt_crossing.Rows[i][col_xing_descr]);
                                            }
                                            string block1 = block_xing_name;
                                            if (dt_crossing.Rows[i][col_xing_block] != DBNull.Value)
                                            {
                                                block1 = Convert.ToString(dt_crossing.Rows[i][col_xing_block]);
                                            }
                                            Colectie_nume_atribute.Add(atr_sta);
                                            Colectie_valori.Add(sta);
                                            Colectie_nume_atribute.Add(atr_descr);
                                            Colectie_valori.Add(Desc1);

                                            Colectie_nume_atribute.Add(atr_sta + "1");
                                            Colectie_valori.Add(sta);
                                            Colectie_nume_atribute.Add(atr_descr + "1");
                                            Colectie_valori.Add(Desc1);

                                            Colectie_nume_atribute.Add(atr_sta + "11");
                                            Colectie_valori.Add(sta);
                                            Colectie_nume_atribute.Add(atr_descr + "11");
                                            Colectie_valori.Add(Desc1);

                                            Colectie_nume_atribute.Add(atr_sta + "111");
                                            Colectie_valori.Add(sta);
                                            Colectie_nume_atribute.Add(atr_descr + "111");
                                            Colectie_valori.Add(Desc1);

                                            Colectie_nume_atribute.Add(atr_sta + "1111");
                                            Colectie_valori.Add(sta);
                                            Colectie_nume_atribute.Add(atr_descr + "1111");
                                            Colectie_valori.Add(Desc1);


                                            BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", block1, Inspt, 1, 0, _AGEN_mainform.layer_crossing_band_text, Colectie_nume_atribute, Colectie_valori);
                                            string visib1 = "";
                                            if (dt_crossing.Rows[i][col_visibility] != DBNull.Value)
                                            {
                                                visib1 = Convert.ToString(dt_crossing.Rows[i][col_visibility]);
                                            }
                                            Functions.set_block_visibility(Block1, visib1);
                                        }
                                        Prevpt = Inspt;
                                    }
                                    rand_prev = nr_rand;
                                }


                                #endregion
                            }
                        }


                        poly3d.Erase();
                        if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
                        {
                            write_crossing_settings_to_excel(_AGEN_mainform.ExcelVisible, _AGEN_mainform.config_path, dt_for_cfg);
                        }
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + debug_i.ToString());
            }

            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();

            this.MdiParent.WindowState = FormWindowState.Normal;
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



        private void checkBox_spit_station_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_split_station.Checked == true)
            {
                checkBox_draw_angle_symbol.Visible = true;
            }
            else
            {
                checkBox_draw_angle_symbol.Checked = false;
                checkBox_draw_angle_symbol.Visible = false;
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

        private void label_crossing_band_Click(object sender, EventArgs e)
        {
            if (panel_dan.Visible == true)
            {
                panel_dan.Visible = false;
            }
            else
            {
                panel_dan.Visible = true;
            }
        }

        private void button_load_tabs_sheet_index_Click(object sender, EventArgs e)
        {
            populate_excel_tabs_to_combobox_from_excel(_AGEN_mainform.sheet_index_excel_name, comboBox_sheet_index_tabs);
        }

        public void populate_excel_tabs_to_combobox_from_excel(string filename, ComboBox combo1)
        {
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {

                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }

                bool excel_is_opened = false;
                string fisier_si = ProjFolder + filename;

                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                Microsoft.Office.Interop.Excel.Application Excel1 = null;

                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");

                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        if (Workbook2.FullName.ToLower() == fisier_si.ToLower())
                        {
                            combo1.Items.Clear();
                            combo1.Items.Add("");

                            Workbook1 = Workbook2;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook2.Worksheets)
                            {
                                combo1.Items.Add(W2.Name);
                            }



                            excel_is_opened = true;
                        }

                    }

                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();

                }


                if (System.IO.File.Exists(fisier_si) == true)
                {
                    if (Workbook1 == null)
                    {
                        if (Excel1.Workbooks.Count == 0) Excel1.Visible = _AGEN_mainform.ExcelVisible;
                        Workbook1 = Excel1.Workbooks.Open(fisier_si);
                    }
                }

                try
                {

                    if (System.IO.File.Exists(fisier_si) == true)
                    {

                        combo1.Items.Clear();
                        combo1.Items.Add("");


                        foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook1.Worksheets)
                        {
                            combo1.Items.Add(W2.Name);
                        }

                        if (excel_is_opened == false)
                        {
                            Workbook1.Close();
                        }


                    }
                    else
                    {
                        combo1.Items.Clear();
                    }



                    if (excel_is_opened == false)
                    {
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
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            else
            {
                MessageBox.Show("the Project database folder location is not specified\r\n" + ProjFolder + "\r\n operation aborted");

                return;
            }
        }

        private void button_load_tabs_crossings_Click(object sender, EventArgs e)
        {
            populate_excel_tabs_to_combobox_from_excel(_AGEN_mainform.crossing_excel_name, comboBox_crossings_tabs);
        }
    }
}
