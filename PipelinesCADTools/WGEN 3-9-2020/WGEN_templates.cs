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
    public partial class Wgen_templates: Form
    {
        int start_row = 2;

        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_create_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_create_all_points);
            lista_butoane.Add(button_all_points_l);
            lista_butoane.Add(button_all_points_nl);
            lista_butoane.Add(button_refresh_ws2);
            lista_butoane.Add(button_create_weld_map);
            lista_butoane.Add(button_weld_map_l);
            lista_butoane.Add(button_weld_map_nl);
            lista_butoane.Add(button_refresh_ws3);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                if (sender as System.Windows.Forms.Button != bt1)
                {
                    bt1.Enabled = false;
                }
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_create_pipe_tally);
            lista_butoane.Add(button_pipe_tally_l);
            lista_butoane.Add(button_pipe_tally_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_create_all_points);
            lista_butoane.Add(button_all_points_l);
            lista_butoane.Add(button_all_points_nl);
            lista_butoane.Add(button_refresh_ws2);
            lista_butoane.Add(button_create_weld_map);
            lista_butoane.Add(button_weld_map_l);
            lista_butoane.Add(button_weld_map_nl);
            lista_butoane.Add(button_refresh_ws3);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_templates()
        {
            InitializeComponent();

            panel_WMT1.Visible = false;
            panel_WM1.Visible = false;

            panel_WMT2.Visible = false;
            panel_WM2.Visible = false;

            panel_APT.Visible = false;
            panel_AP.Visible = false;

            panel_PTT.Visible = false;
            panel_PT.Visible = false;
        }

        private void button_create_pipe_tally_Click(object sender, EventArgs e)
        {
            string col1 = "MMID";
            string col2 = "Pipe";
            string col3 = "Heat";
            string col4 = "OriginalLength";
            string col5 = "NewLength";
            string col6 = "WallThickness";
            string col7 = "Diameter";
            string col8 = "Grade";
            string col9 = "Coating";
            string col10 = "Manufacture";
            string col11 = "DoubleJointNo";
            Wgen_main_form.dt_ground_tally = Functions.Creaza_weldmap_pipe_tally_datatable_structure();
            string colpt1 = textBox_pt1.Text;
            string colpt2 = textBox_pt2.Text;
            string colpt3 = textBox_pt3.Text;
            string colpt4 = textBox_pt4.Text;
            string colpt5 = textBox_pt5.Text;
            string colpt6 = textBox_pt6.Text;
            string colpt7 = textBox_pt7.Text;
            string colpt8 = textBox_pt8.Text;
            string colpt9 = textBox_pt9.Text;
            string colpt10 = textBox_pt10.Text;
            string colpt11 = textBox_pt11.Text;

            if (comboBox_ws1.Text != "")
            {
                string string1 = comboBox_ws1.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false(sender);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Wgen_main_form.dt_ground_tally = Functions.build_data_table_from_excel_based_on_columns(Wgen_main_form.dt_ground_tally, W1, start_row,
                                                                col1, colpt1, col2, colpt2, col3, colpt3, col4, colpt4, col5, colpt5, col6, colpt6, col7, colpt7,
                                                                col8, colpt8, col9, colpt9, col10, colpt10, col11, colpt11, "", "", "", "", "", "", "", "", "", "", "", ""
                                                                , "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
                        }
                    }
                }
            }
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Wgen_main_form.dt_ground_tally);
            set_enable_true();
            button_pipe_tally_l.Visible = true;
            button_pipe_tally_nl.Visible = false;
        }

        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            if (comboBox_ws1.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
                {
                    if (comboBox_ws1.Items[i].ToString().ToUpper().Contains("GROUND_TALLY") == true)
                    {
                        comboBox_ws1.SelectedIndex = i;
                        i = comboBox_ws1.Items.Count;
                    }
                }
            }
        }


        private void comboBox_ws1_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel_PTT.Visible = true;
            panel_PT.Visible = true;
            if (panel_WMT1.Visible == false && panel_APT.Visible == false)
            {
                panel_PTT.Location = new Point(3, 188);
                panel_PT.Location = new Point(3, 213);
            }
            else if (panel_WMT1.Visible == false && panel_APT.Visible == true)
            {
                panel_PTT.Location = new Point(249, 188);
                panel_PT.Location = new Point(249, 213);
                panel_APT.Location = new Point(3, 188);
                panel_AP.Location = new Point(3, 213);
            }
            else
            {
                panel_PTT.Location = new Point(249, 315);
                panel_PT.Location = new Point(249, 341);
                panel_WMT1.Location = new Point(3, 188);
                panel_WM1.Location = new Point(3, 213);
                panel_WMT2.Location = new Point(249, 188);
                panel_WM2.Location = new Point(249, 213);
                panel_APT.Location = new Point(463, 188);
                panel_AP.Location = new Point(463, 213);
            }
        }

        private void button_refresh_ws2_Click(object sender, EventArgs e)
        {

            Functions.Load_opened_worksheets_to_combobox(comboBox_ws2);
            if (comboBox_ws2.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws2.Items.Count; ++i)
                {
                    if (comboBox_ws2.Items[i].ToString().ToUpper().Contains("ALL_POINTS") == true)
                    {
                        comboBox_ws2.SelectedIndex = i;
                        i = comboBox_ws2.Items.Count;
                    }
                }
            }
        }

        private void comboBox_ws2_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel_APT.Visible = true;
            panel_AP.Visible = true;
            if (panel_WMT1.Visible == false && panel_PTT.Visible == false)
            {
                panel_APT.Location = new Point(3, 188);
                panel_AP.Location = new Point(3, 213);
            }
            else if (panel_WMT1.Visible == false && panel_PTT.Visible == true)
            {
                panel_APT.Location = new Point(3, 188);
                panel_AP.Location = new Point(3, 213);
                panel_PTT.Location = new Point(249, 188);
                panel_PT.Location = new Point(249, 213);
            }
            else
            {
                panel_APT.Location = new Point(463, 188);
                panel_AP.Location = new Point(463, 213);
                panel_PTT.Location = new Point(249, 315);
                panel_PT.Location = new Point(249, 341);
                panel_WMT1.Location = new Point(3, 188);
                panel_WM1.Location = new Point(3, 213);
                panel_WMT2.Location = new Point(249, 188);
                panel_WM2.Location = new Point(249, 213);
            }
        }


        private void comboBox_ws3_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel_WMT1.Visible = true;
            panel_WM1.Visible = true;
            panel_WMT2.Visible = true;
            panel_WM2.Visible = true;

            panel_APT.Location = new Point(463, 188);
            panel_AP.Location = new Point(463, 213);
            panel_PTT.Location = new Point(249, 315);
            panel_PT.Location = new Point(249, 341);
            panel_WMT1.Location = new Point(3, 188);
            panel_WM1.Location = new Point(3, 213);
            panel_WMT2.Location = new Point(249, 188);
            panel_WM2.Location = new Point(249, 213);

        }

        private void button_refresh_ws3_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws3);
            if (comboBox_ws3.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws3.Items.Count; ++i)
                {
                    if (comboBox_ws3.Items[i].ToString().ToUpper().Contains("WELD_MAP") == true)
                    {
                        comboBox_ws3.SelectedIndex = i;
                        i = comboBox_ws3.Items.Count;
                    }
                }
            }
        }

        private void button_create_all_points_Click(object sender, EventArgs e)
        {
            string col1 = "PNT";
            string col2 = "NORTHING";
            string col3 = "EASTING";
            string col4 = "ELEVATION";
            string col5 = "FEATURE CODE";
            string col6 = "STATION";
            string col7 = "FILENAME";
            string col8 = "LOCATION";
            string col9 = "NOTES";
            string col10 = "DESCRIPTION";
            string col11 = "MISC1";
            string col12 = "H_ANGLE";
            string col13 = "V_ANGLE";
            string col14 = "MISC4";
            string col15 = "MISC5";
            string col16 = "MISC6";
            string col17 = "MISC7";

            Wgen_main_form.dt_all_points = Functions.Creaza_all_points_datatable_structure();
            string colpt1 = textBox_ap1.Text;
            string colpt2 = textBox_ap2.Text;
            string colpt3 = textBox_ap3.Text;
            string colpt4 = textBox_ap4.Text;
            string colpt5 = textBox_ap5.Text;
            string colpt6 = textBox_ap6.Text;
            string colpt7 = textBox_ap7.Text;
            string colpt8 = textBox_ap8.Text;
            string colpt9 = textBox_ap9.Text;
            string colpt10 = textBox_ap10.Text;
            string colpt11 = textBox_ap11.Text;
            string colpt12 = textBox_ap12.Text;
            string colpt13 = textBox_ap13.Text;
            string colpt14 = textBox_ap14.Text;
            string colpt15 = textBox_ap15.Text;
            string colpt16 = textBox_ap16.Text;
            string colpt17 = textBox_ap17.Text;

            if (comboBox_ws2.Text != "")
            {
                string string1 = comboBox_ws2.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false(sender);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Wgen_main_form.dt_all_points = Functions.build_data_table_from_excel_based_on_columns(Wgen_main_form.dt_all_points, W1, start_row,
                                                                col1, colpt1, col2, colpt2, col3, colpt3, col4, colpt4, col5, colpt5, col6, colpt6,
                                                                col7, colpt7, col8, colpt8, col9, colpt9, col10, colpt10, col11, colpt11,
                                                                col12, colpt12, col13, colpt13, col14, colpt14, col15, colpt15, col16, colpt16, col17, colpt17,
                                                                 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
                        }
                    }
                }
            }
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Wgen_main_form.dt_all_points);
            set_enable_true();
            button_all_points_l.Visible = true;
            button_all_points_nl.Visible = false;
        }

        private void button_create_weld_map_Click(object sender, EventArgs e)
        {
            string col1 = "PNT";
            string col2 = "NORTHING";
            string col3 = "EASTING";
            string col4 = "ELEVATION";
            string col5 = "FEATURE_CODE";
            string col6 = "DESCRIPTION";
            string col7 = "PROJECT_STATION";
            string col8 = "MM_BK";
            string col9 = "WALL_BK";
            string col10 = "PIPE_BK";
            string col11 = "HEAT_BK";
            string col12 = "COATING_BK";
            string col13 = "MM_AHD";
            string col14 = "WALL_AHD";
            string col15 = "PIPE_AHD";
            string col16 = "HEAT_AHD";
            string col17 = "COATING_AHD";
            string col18 = "NG";
            string col19 = "NG_NORTHING";
            string col20 = "NG_EASTING";
            string col21 = "NG_ELEVATION";
            string col22 = "COVER";
            string col23 = "LOCATION";
            string col24 = "FILENAME";
            string col25 = "H_ANGLE";
            string col26 = "V_ANGLE";
            string col27 = "CROSSING_NAME";

            Wgen_main_form.dt_weld_map = Functions.Creaza_weldmap_datatable_structure();
            string colpt1 = textBox_wm1.Text;
            string colpt2 = textBox_wm2.Text;
            string colpt3 = textBox_wm3.Text;
            string colpt4 = textBox_wm4.Text;
            string colpt5 = textBox_wm5.Text;
            string colpt6 = textBox_wm6.Text;
            string colpt7 = textBox_wm7.Text;
            string colpt8 = textBox_wm8.Text;
            string colpt9 = textBox_wm9.Text;
            string colpt10 = textBox_wm10.Text;
            string colpt11 = textBox_wm11.Text;
            string colpt12 = textBox_wm12.Text;
            string colpt13 = textBox_wm13.Text;
            string colpt14 = textBox_wm14.Text;
            string colpt15 = textBox_wm15.Text;
            string colpt16 = textBox_wm16.Text;
            string colpt17 = textBox_wm17.Text;
            string colpt18 = textBox_wm18.Text;
            string colpt19 = textBox_wm19.Text;
            string colpt20 = textBox_wm20.Text;
            string colpt21 = textBox_wm21.Text;
            string colpt22 = textBox_wm22.Text;
            string colpt23 = textBox_wm23.Text;
            string colpt24 = textBox_wm24.Text;
            string colpt25 = textBox_wm25.Text;
            string colpt26 = textBox_wm26.Text;
            string colpt27 = textBox_wm27.Text;

            if (comboBox_ws3.Text != "")
            {
                string string1 = comboBox_ws3.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false(sender);
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        if (W1 != null)
                        {
                            Wgen_main_form.dt_weld_map = Functions.build_data_table_from_excel_based_on_columns(Wgen_main_form.dt_all_points, W1, start_row,
                                                                col1, colpt1, col2, colpt2, col3, colpt3, col4, colpt4, col5, colpt5, col6, colpt6,
                                                                col7, colpt7, col8, colpt8, col9, colpt9, col10, colpt10, col11, colpt11,
                                                                col12, colpt12, col13, colpt13, col14, colpt14, col15, colpt15, col16, colpt16,
                                                                col17, colpt17, col18, colpt18, col19, colpt19, col20, colpt20, col21, colpt21,
                                                                col22, colpt22, col23, colpt23, col24, colpt24, col25, colpt25, col26, colpt26, col27, colpt27);
                        }
                    }
                }
            }
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Wgen_main_form.dt_weld_map);
            set_enable_true();
            button_weld_map_l.Visible = true;
            button_weld_map_nl.Visible = false;
        }
    }
}
