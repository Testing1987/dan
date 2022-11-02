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
using Font = System.Drawing.Font;

namespace Alignment_mdi
{
    public partial class Wgen_all_pts : Form
    {

        private ContextMenuStrip ContextMenuStrip_go_to_error;

        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_export;
        int start_row = 2;

        System.Data.DataTable dt_display;

        public System.Data.DataTable dt_dismissed_errors = null;
        public string dismiss_errors_tab = "Dsd errors ap";
        Microsoft.Office.Interop.Excel.Worksheet W2 = null;
        public string filename = "";

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_load_all_points);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }


        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_load_all_points);
            lista_butoane.Add(button_all_pts_l);
            lista_butoane.Add(button_all_pts_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_export_errors_to_xl);

            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_all_pts()
        {
            InitializeComponent();

            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Go to error" };
            toolStripMenuItem1.Click += go_to_excel_point;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Zoom to point in AutoCAD" };
            toolStripMenuItem2.Click += zoom_to_point_in_acad;

            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2 });

            if (Wgen_main_form.client_name == "xxx") Wgen_main_form.client_name = Wgen_main_form.lista_clienti[0];
            label_client.Text = Wgen_main_form.client_name;

        }

        public void set_label_client(string continut)
        {
            label_client.Text = continut;
        }

        private void button_load_all_pts_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_all_points = Functions.Creaza_all_points_datatable_structure();

            make_first_line_invisible();
            if (comboBox_ws1.Text != "")
            {
                string string1 = comboBox_ws1.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                     filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);

                        W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, dismiss_errors_tab);

                        if (W1 != null)
                        {


                            if (W2 != null)
                            {
                                dt_dismissed_errors = new System.Data.DataTable();
                                dt_dismissed_errors.Columns.Add("Point", typeof(string));
                                dt_dismissed_errors.Columns.Add("Value", typeof(string));
                                dt_dismissed_errors.Columns.Add("Error", typeof(string));

                                dt_dismissed_errors = Functions.Populate_data_table_from_excel(dt_dismissed_errors, W2, start_row,
                                    "A", "B", "C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", false);
                                if (dt_dismissed_errors.Rows.Count == 0) dt_dismissed_errors = null;
                            }
                            else
                            {
                                dt_dismissed_errors = null;
                            }

                            Wgen_main_form.dt_weld_map = null;
                            Wgen_main_form.dt_pt_keep = null;
                            Wgen_main_form.dt_pt_move = null;

                            System.Data.DataTable dt1 = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_all_points, W1, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text,
                                textBox_6.Text, textBox_7.Text, textBox_8.Text, textBox_9.Text, textBox_10.Text, textBox_11.Text, textBox_12.Text, textBox_13.Text, textBox_14.Text, textBox_15.Text, textBox_16.Text,
                                textBox_17.Text, textBox_18.Text, textBox_19.Text, textBox_20.Text, textBox_21.Text, textBox_22.Text, textBox_23.Text, textBox_24.Text, textBox_25.Text, textBox_26.Text, true);
                            dt1.Columns.Add("index1", typeof(int));
                            Wgen_main_form.dt_all_points = dt1.Clone();



                            Wgen_main_form.dt_all_points.TableName = "TABLA_ALLPT";
                            if (dt1.Rows.Count > 0)
                            {



                                int nr_duplicates = 0;
                                int nr_mmid_not_found = 0;
                                int nr_xray_duplicates = 0;
                                int nr_null_values = 0;
                                dt_errors = new System.Data.DataTable();
                                dt_errors.Columns.Add("Point", typeof(string));
                                dt_errors.Columns.Add("Value", typeof(string));
                                dt_errors.Columns.Add("Excel", typeof(string));
                                dt_errors.Columns.Add("w1", typeof(Microsoft.Office.Interop.Excel.Worksheet));
                                dt_errors.Columns.Add("Error", typeof(string));
                                dt_errors.Columns.Add("x", typeof(string));
                                dt_errors.Columns.Add("y", typeof(string));

                                string col1 = "PNT";
                                string col2 = "NORTHING";
                                string col3 = "EASTING";
                                string col4 = "ELEVATION";
                                string col5 = "FEATURE CODE";

                                string col6 = "FILENAME";
                                string col7 = "LOCATION";
                                string col8 = "NOTES";
                                string col9 = "DESCRIPTION";
                                string col10 = "MISC1";
                                string col11 = "MISC2";
                                string col12 = "MISC3";
                                string col13 = "MISC4";
                                string col14 = "MISC5";
                                string col15 = "MISC6";
                                string col16 = "MISC7";
                                string col17 = "MISC8";
                                string col18 = "MISC9";
                                string col19 = "MISC10";
                                string col20 = "MISC11";
                                string col21 = "MISC12";
                                string col22 = "MISC13";
                                string col23 = "MISC14";
                                string col24 = "MISC15";
                                string col25 = "MISC16";
                                string col26 = "MISC17";

                                string colGT1 = "MMID";
                                string colGT2 = "Pipe";
                                string colGT3 = "Heat";
                                string colGT4 = "OriginalLength";
                                string colGT5 = "NewLength";
                                string colGT6 = "WallThickness";
                                string colGT7 = "Diameter";
                                string colGT8 = "Grade";
                                string colGT9 = "Coating";
                                string colGT10 = "Manufacture";



                                for (int i = 0; i < dt1.Rows.Count; ++i)
                                {
                                    dt1.Rows[i]["index1"] = i;

                                    Wgen_main_form.dt_all_points.ImportRow(dt1.Rows[i]);
                                }

                                string col_xray = col8;
                                string col_bend_type = col7;
                                string col_bend_defl = col8;
                                string col_bend_pos = col9;
                                string col_bend_hor = col10;
                                string col_bend_ver = col11;

                                string col_mm_back = col11;
                                string col_mm_ahead = col12;

                                string xl_xray = "A";
                                string xl_bend_type = "A";
                                string xl_bend_defl = "A";
                                string xl_bend_pos = "A";
                                string xl_bend_hor = "A";
                                string xl_bend_ver = "A";

                                string xl_elbow_type = "A";
                                string xl_elbow_defl = "A";
                                string xl_elbow_pos = "A";
                                string xl_elbow_hor = "A";
                                string xl_elbow_ver = "A";
                                string col_elbow_type = col7;
                                string col_elbow_defl = col8;
                                string col_elbow_pos = col9;
                                string col_elbow_hor = col10;
                                string col_elbow_ver = col11;


                                string xl_mm_back = "A";
                                string xl_mm_ahead = "A";


                                for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                                {
                                    string client_name = "xxx";
                                    if (Wgen_main_form.dt_feature_codes.Rows[j][0] != DBNull.Value)
                                    {
                                        client_name = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]);
                                        if (client_name == Wgen_main_form.client_name)
                                        {
                                            if (Wgen_main_form.dt_feature_codes.Rows[j][1] != DBNull.Value)
                                            {
                                                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "WELD" || Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "WLD")
                                                {
                                                    #region XRAY
                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][14] != DBNull.Value)
                                                    {
                                                        string check1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][14]);
                                                        if (check1.Contains("{F}") == true ||
                                                            check1.Contains("{G}") == true ||
                                                            check1.Contains("{H}") == true ||
                                                            check1.Contains("{I}") == true ||
                                                            check1.Contains("{J}") == true ||
                                                            check1.Contains("{K}") == true ||
                                                            check1.Contains("{L}") == true ||
                                                            check1.Contains("{M}") == true ||
                                                            check1.Contains("{N}") == true ||
                                                            check1.Contains("{O}") == true ||
                                                            check1.Contains("{P}") == true
                                                            )
                                                        {
                                                            if (check1.Contains("{F}") == true)
                                                            {
                                                                col_xray = col6;
                                                                xl_xray = "F";

                                                            }
                                                            if (check1.Contains("{G}") == true)
                                                            {
                                                                col_xray = col7;
                                                                xl_xray = "G";

                                                            }
                                                            if (check1.Contains("{H}") == true)
                                                            {
                                                                col_xray = col8;
                                                                xl_xray = "H";

                                                            }
                                                            if (check1.Contains("{I}") == true)
                                                            {
                                                                col_xray = col9;
                                                                xl_xray = "I";

                                                            }
                                                            if (check1.Contains("{J}") == true)
                                                            {
                                                                col_xray = col10;
                                                                xl_xray = "J";

                                                            }
                                                            if (check1.Contains("{K}") == true)
                                                            {
                                                                col_xray = col11;
                                                                xl_xray = "K";

                                                            }
                                                            if (check1.Contains("{L}") == true)
                                                            {
                                                                col_xray = col12;
                                                                xl_xray = "L";

                                                            }
                                                            if (check1.Contains("{M}") == true)
                                                            {
                                                                col_xray = col13;
                                                                xl_xray = "M";

                                                            }
                                                            if (check1.Contains("{N}") == true)
                                                            {
                                                                col_xray = col14;
                                                                xl_xray = "N";

                                                            }
                                                            if (check1.Contains("{O}") == true)
                                                            {
                                                                col_xray = col15;
                                                                xl_xray = "O";

                                                            }
                                                            if (check1.Contains("{P}") == true)
                                                            {
                                                                col_xray = col16;
                                                                xl_xray = "P";

                                                            }




                                                        }
                                                    }
                                                    #endregion                                                  

                                                    #region mm back
                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][20] != DBNull.Value)
                                                    {
                                                        string check1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][20]);
                                                        if (check1.Contains("{F}") == true ||
                                                            check1.Contains("{G}") == true ||
                                                            check1.Contains("{H}") == true ||
                                                            check1.Contains("{I}") == true ||
                                                            check1.Contains("{J}") == true ||
                                                            check1.Contains("{K}") == true ||
                                                            check1.Contains("{L}") == true ||
                                                            check1.Contains("{M}") == true ||
                                                            check1.Contains("{N}") == true ||
                                                            check1.Contains("{O}") == true ||
                                                            check1.Contains("{P}") == true
                                                            )
                                                        {
                                                            if (check1.Contains("{F}") == true)
                                                            {
                                                                col_mm_back = col6;
                                                                xl_mm_back = "F";

                                                            }
                                                            if (check1.Contains("{G}") == true)
                                                            {
                                                                col_mm_back = col7;
                                                                xl_mm_back = "G";
                                                            }
                                                            if (check1.Contains("{H}") == true)
                                                            {
                                                                col_mm_back = col8;
                                                                xl_mm_back = "H";

                                                            }
                                                            if (check1.Contains("{I}") == true)
                                                            {
                                                                col_mm_back = col9;
                                                                xl_mm_back = "I";

                                                            }
                                                            if (check1.Contains("{J}") == true)
                                                            {
                                                                col_mm_back = col10;
                                                                xl_mm_back = "J";

                                                            }
                                                            if (check1.Contains("{K}") == true)
                                                            {
                                                                col_mm_back = col11;
                                                                xl_mm_back = "K";

                                                            }
                                                            if (check1.Contains("{L}") == true)
                                                            {
                                                                col_mm_back = col12;
                                                                xl_mm_back = "L";

                                                            }
                                                            if (check1.Contains("{M}") == true)
                                                            {
                                                                col_mm_back = col13;
                                                                xl_mm_back = "M";

                                                            }
                                                            if (check1.Contains("{N}") == true)
                                                            {
                                                                col_mm_back = col14;
                                                                xl_mm_back = "N";

                                                            }
                                                            if (check1.Contains("{O}") == true)
                                                            {
                                                                col_mm_back = col15;
                                                                xl_mm_back = "O";

                                                            }
                                                            if (check1.Contains("{P}") == true)
                                                            {
                                                                col_mm_back = col16;
                                                                xl_mm_back = "P";

                                                            }
                                                        }
                                                    }
                                                    #endregion

                                                    #region mm ahead
                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][21] != DBNull.Value)
                                                    {
                                                        string check1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][21]);
                                                        if (check1.Contains("{F}") == true ||
                                                            check1.Contains("{G}") == true ||
                                                            check1.Contains("{H}") == true ||
                                                            check1.Contains("{I}") == true ||
                                                            check1.Contains("{J}") == true ||
                                                            check1.Contains("{K}") == true ||
                                                            check1.Contains("{L}") == true ||
                                                            check1.Contains("{M}") == true ||
                                                            check1.Contains("{N}") == true ||
                                                            check1.Contains("{O}") == true ||
                                                            check1.Contains("{P}") == true
                                                            )
                                                        {
                                                            if (check1.Contains("{F}") == true)
                                                            {
                                                                col_mm_ahead = col6;
                                                                xl_mm_ahead = "F";

                                                            }
                                                            if (check1.Contains("{G}") == true)
                                                            {
                                                                col_mm_ahead = col7;
                                                                xl_mm_ahead = "G";

                                                            }
                                                            if (check1.Contains("{H}") == true)
                                                            {
                                                                col_mm_ahead = col8;
                                                                xl_mm_ahead = "H";

                                                            }
                                                            if (check1.Contains("{I}") == true)
                                                            {
                                                                col_mm_ahead = col9;
                                                                xl_mm_ahead = "I";

                                                            }
                                                            if (check1.Contains("{J}") == true)
                                                            {
                                                                col_mm_ahead = col10;
                                                                xl_mm_ahead = "J";

                                                            }
                                                            if (check1.Contains("{K}") == true)
                                                            {
                                                                col_mm_ahead = col11;
                                                                xl_mm_ahead = "K";

                                                            }
                                                            if (check1.Contains("{L}") == true)
                                                            {
                                                                col_mm_ahead = col12;
                                                                xl_mm_ahead = "L";

                                                            }
                                                            if (check1.Contains("{M}") == true)
                                                            {
                                                                col_mm_ahead = col13;
                                                                xl_mm_ahead = "M";

                                                            }
                                                            if (check1.Contains("{N}") == true)
                                                            {
                                                                col_mm_ahead = col14;
                                                                xl_mm_ahead = "N";

                                                            }
                                                            if (check1.Contains("{O}") == true)
                                                            {
                                                                col_mm_ahead = col15;
                                                                xl_mm_ahead = "O";

                                                            }
                                                            if (check1.Contains("{P}") == true)
                                                            {
                                                                col_mm_ahead = col16;
                                                                xl_mm_ahead = "P";

                                                            }




                                                        }
                                                    }
                                                    #endregion                                                  
                                                }

                                                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "BEND")
                                                {
                                                    #region BEND
                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][15] != DBNull.Value)
                                                    {
                                                        #region BEND TYPE
                                                        string check15 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][15]);
                                                        if (check15.Contains("{F}") == true ||
                                                            check15.Contains("{G}") == true ||
                                                            check15.Contains("{H}") == true ||
                                                            check15.Contains("{I}") == true ||
                                                            check15.Contains("{J}") == true ||
                                                            check15.Contains("{K}") == true ||
                                                            check15.Contains("{L}") == true ||
                                                            check15.Contains("{M}") == true ||
                                                            check15.Contains("{N}") == true ||
                                                            check15.Contains("{O}") == true ||
                                                            check15.Contains("{P}") == true ||
                                                            check15.Contains("{Q}") == true ||
                                                            check15.Contains("{R}") == true ||
                                                            check15.Contains("{S}") == true ||
                                                            check15.Contains("{T}") == true ||
                                                            check15.Contains("{U}") == true ||
                                                            check15.Contains("{V}") == true ||
                                                            check15.Contains("{W}") == true ||
                                                            check15.Contains("{X}") == true ||
                                                            check15.Contains("{Y}") == true ||
                                                            check15.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check15.Contains("{F}") == true)
                                                            {
                                                                col_bend_type = col6;
                                                                xl_bend_type = "F";

                                                            }
                                                            if (check15.Contains("{G}") == true)
                                                            {
                                                                col_bend_type = col7;
                                                                xl_bend_type = "G";

                                                            }
                                                            if (check15.Contains("{H}") == true)
                                                            {
                                                                col_bend_type = col8;
                                                                xl_bend_type = "H";

                                                            }
                                                            if (check15.Contains("{I}") == true)
                                                            {
                                                                col_bend_type = col9;
                                                                xl_bend_type = "I";

                                                            }
                                                            if (check15.Contains("{J}") == true)
                                                            {
                                                                col_bend_type = col10;
                                                                xl_bend_type = "J";

                                                            }
                                                            if (check15.Contains("{K}") == true)
                                                            {
                                                                col_bend_type = col11;
                                                                xl_bend_type = "K";

                                                            }
                                                            if (check15.Contains("{L}") == true)
                                                            {
                                                                col_bend_type = col12;
                                                                xl_bend_type = "L";

                                                            }
                                                            if (check15.Contains("{M}") == true)
                                                            {
                                                                col_bend_type = col13;
                                                                xl_bend_type = "M";

                                                            }
                                                            if (check15.Contains("{N}") == true)
                                                            {
                                                                col_bend_type = col14;
                                                                xl_bend_type = "N";

                                                            }
                                                            if (check15.Contains("{O}") == true)
                                                            {
                                                                col_bend_type = col15;
                                                                xl_bend_type = "O";

                                                            }
                                                            if (check15.Contains("{P}") == true)
                                                            {
                                                                col_bend_type = col16;
                                                                xl_bend_type = "P";

                                                            }
                                                            if (check15.Contains("{Q}") == true)
                                                            {
                                                                col_bend_type = col17;
                                                                xl_bend_type = "Q";

                                                            }
                                                            if (check15.Contains("{R}") == true)
                                                            {
                                                                col_bend_type = col18;
                                                                xl_bend_type = "R";

                                                            }
                                                            if (check15.Contains("{S}") == true)
                                                            {
                                                                col_bend_type = col19;
                                                                xl_bend_type = "S";

                                                            }
                                                            if (check15.Contains("{T}") == true)
                                                            {
                                                                col_bend_type = col20;
                                                                xl_bend_type = "T";

                                                            }
                                                            if (check15.Contains("{U}") == true)
                                                            {
                                                                col_bend_type = col21;
                                                                xl_bend_type = "U";

                                                            }
                                                            if (check15.Contains("{V}") == true)
                                                            {
                                                                col_bend_type = col22;
                                                                xl_bend_type = "V";

                                                            }
                                                            if (check15.Contains("{W}") == true)
                                                            {
                                                                col_bend_type = col23;
                                                                xl_bend_type = "W";

                                                            }
                                                            if (check15.Contains("{X}") == true)
                                                            {
                                                                col_bend_type = col24;
                                                                xl_bend_type = "X";

                                                            }
                                                            if (check15.Contains("{Y}") == true)
                                                            {
                                                                col_bend_type = col25;
                                                                xl_bend_type = "Y";

                                                            }
                                                            if (check15.Contains("{Z}") == true)
                                                            {
                                                                col_bend_type = col26;
                                                                xl_bend_type = "Z";

                                                            }



                                                        }
                                                        #endregion
                                                    }

                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][16] != DBNull.Value)
                                                    {
                                                        #region BEND DEFLECTION
                                                        string check16 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][16]);
                                                        if (check16.Contains("{F}") == true ||
                                                            check16.Contains("{G}") == true ||
                                                            check16.Contains("{H}") == true ||
                                                            check16.Contains("{I}") == true ||
                                                            check16.Contains("{J}") == true ||
                                                            check16.Contains("{K}") == true ||
                                                            check16.Contains("{L}") == true ||
                                                            check16.Contains("{M}") == true ||
                                                            check16.Contains("{N}") == true ||
                                                            check16.Contains("{O}") == true ||
                                                            check16.Contains("{P}") == true ||
                                                            check16.Contains("{Q}") == true ||
                                                            check16.Contains("{R}") == true ||
                                                            check16.Contains("{S}") == true ||
                                                            check16.Contains("{T}") == true ||
                                                            check16.Contains("{U}") == true ||
                                                            check16.Contains("{V}") == true ||
                                                            check16.Contains("{W}") == true ||
                                                            check16.Contains("{X}") == true ||
                                                            check16.Contains("{Y}") == true ||
                                                            check16.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check16.Contains("{F}") == true)
                                                            {
                                                                col_bend_defl = col6;
                                                                xl_bend_defl = "F";

                                                            }
                                                            if (check16.Contains("{G}") == true)
                                                            {
                                                                col_bend_defl = col7;
                                                                xl_bend_defl = "G";

                                                            }
                                                            if (check16.Contains("{H}") == true)
                                                            {
                                                                col_bend_defl = col8;
                                                                xl_bend_defl = "H";

                                                            }
                                                            if (check16.Contains("{I}") == true)
                                                            {
                                                                col_bend_defl = col9;
                                                                xl_bend_defl = "I";

                                                            }
                                                            if (check16.Contains("{J}") == true)
                                                            {
                                                                col_bend_defl = col10;
                                                                xl_bend_defl = "J";

                                                            }
                                                            if (check16.Contains("{K}") == true)
                                                            {
                                                                col_bend_defl = col11;
                                                                xl_bend_defl = "K";

                                                            }
                                                            if (check16.Contains("{L}") == true)
                                                            {
                                                                col_bend_defl = col12;
                                                                xl_bend_defl = "L";

                                                            }
                                                            if (check16.Contains("{M}") == true)
                                                            {
                                                                col_bend_defl = col13;
                                                                xl_bend_defl = "M";

                                                            }
                                                            if (check16.Contains("{N}") == true)
                                                            {
                                                                col_bend_defl = col14;
                                                                xl_bend_defl = "N";

                                                            }
                                                            if (check16.Contains("{O}") == true)
                                                            {
                                                                col_bend_defl = col15;
                                                                xl_bend_defl = "O";

                                                            }
                                                            if (check16.Contains("{P}") == true)
                                                            {
                                                                col_bend_defl = col16;
                                                                xl_bend_defl = "P";

                                                            }

                                                            if (check16.Contains("{Q}") == true)
                                                            {
                                                                col_bend_defl = col17;
                                                                xl_bend_defl = "Q";

                                                            }
                                                            if (check16.Contains("{R}") == true)
                                                            {
                                                                col_bend_defl = col18;
                                                                xl_bend_defl = "R";

                                                            }
                                                            if (check16.Contains("{S}") == true)
                                                            {
                                                                col_bend_defl = col19;
                                                                xl_bend_defl = "S";

                                                            }
                                                            if (check16.Contains("{T}") == true)
                                                            {
                                                                col_bend_defl = col20;
                                                                xl_bend_defl = "T";

                                                            }
                                                            if (check16.Contains("{U}") == true)
                                                            {
                                                                col_bend_defl = col21;
                                                                xl_bend_defl = "U";

                                                            }
                                                            if (check16.Contains("{V}") == true)
                                                            {
                                                                col_bend_defl = col22;
                                                                xl_bend_defl = "V";

                                                            }
                                                            if (check16.Contains("{W}") == true)
                                                            {
                                                                col_bend_defl = col23;
                                                                xl_bend_defl = "W";

                                                            }
                                                            if (check16.Contains("{X}") == true)
                                                            {
                                                                col_bend_defl = col24;
                                                                xl_bend_defl = "X";

                                                            }
                                                            if (check16.Contains("{Y}") == true)
                                                            {
                                                                col_bend_defl = col25;
                                                                xl_bend_defl = "Y";

                                                            }
                                                            if (check16.Contains("{Z}") == true)
                                                            {
                                                                col_bend_defl = col26;
                                                                xl_bend_defl = "Z";

                                                            }


                                                        }
                                                        #endregion
                                                    }


                                                  

                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][17] != DBNull.Value)
                                                    {
                                                        #region BEND POSITION
                                                        string check17 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][17]);
                                                        if (check17.Contains("{F}") == true ||
                                                            check17.Contains("{G}") == true ||
                                                            check17.Contains("{H}") == true ||
                                                            check17.Contains("{I}") == true ||
                                                            check17.Contains("{J}") == true ||
                                                            check17.Contains("{K}") == true ||
                                                            check17.Contains("{L}") == true ||
                                                            check17.Contains("{M}") == true ||
                                                            check17.Contains("{N}") == true ||
                                                            check17.Contains("{O}") == true ||
                                                            check17.Contains("{P}") == true ||
                                                            check17.Contains("{Q}") == true ||
                                                            check17.Contains("{R}") == true ||
                                                            check17.Contains("{S}") == true ||
                                                            check17.Contains("{T}") == true ||
                                                            check17.Contains("{U}") == true ||
                                                            check17.Contains("{V}") == true ||
                                                            check17.Contains("{W}") == true ||
                                                            check17.Contains("{X}") == true ||
                                                            check17.Contains("{Y}") == true ||
                                                            check17.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check17.Contains("{F}") == true)
                                                            {
                                                                col_bend_pos = col6;
                                                                xl_bend_pos = "F";

                                                            }
                                                            if (check17.Contains("{G}") == true)
                                                            {
                                                                col_bend_pos = col7;
                                                                xl_bend_pos = "G";

                                                            }
                                                            if (check17.Contains("{H}") == true)
                                                            {
                                                                col_bend_pos = col8;
                                                                xl_bend_pos = "H";

                                                            }
                                                            if (check17.Contains("{I}") == true)
                                                            {
                                                                col_bend_pos = col9;
                                                                xl_bend_pos = "I";

                                                            }
                                                            if (check17.Contains("{J}") == true)
                                                            {
                                                                col_bend_pos = col10;
                                                                xl_bend_pos = "J";

                                                            }
                                                            if (check17.Contains("{K}") == true)
                                                            {
                                                                col_bend_pos = col11;
                                                                xl_bend_pos = "K";

                                                            }
                                                            if (check17.Contains("{L}") == true)
                                                            {
                                                                col_bend_pos = col12;
                                                                xl_bend_pos = "L";

                                                            }
                                                            if (check17.Contains("{M}") == true)
                                                            {
                                                                col_bend_pos = col13;

                                                            }
                                                            if (check17.Contains("{N}") == true)
                                                            {
                                                                col_bend_pos = col14;
                                                                xl_bend_pos = "N";

                                                            }
                                                            if (check17.Contains("{O}") == true)
                                                            {
                                                                col_bend_pos = col15;
                                                                xl_bend_pos = "O";

                                                            }
                                                            if (check17.Contains("{P}") == true)
                                                            {
                                                                col_bend_pos = col16;
                                                                xl_bend_pos = "P";

                                                            }
                                                            if (check17.Contains("{Q}") == true)
                                                            {
                                                                col_bend_pos = col17;
                                                                xl_bend_pos = "Q";

                                                            }
                                                            if (check17.Contains("{R}") == true)
                                                            {
                                                                col_bend_pos = col18;
                                                                xl_bend_pos = "R";

                                                            }
                                                            if (check17.Contains("{S}") == true)
                                                            {
                                                                col_bend_pos = col19;
                                                                xl_bend_pos = "S";

                                                            }
                                                            if (check17.Contains("{T}") == true)
                                                            {
                                                                col_bend_pos = col20;
                                                                xl_bend_pos = "T";

                                                            }
                                                            if (check17.Contains("{U}") == true)
                                                            {
                                                                col_bend_pos = col21;
                                                                xl_bend_pos = "U";

                                                            }
                                                            if (check17.Contains("{V}") == true)
                                                            {
                                                                col_bend_pos = col22;
                                                                xl_bend_pos = "V";

                                                            }
                                                            if (check17.Contains("{W}") == true)
                                                            {
                                                                col_bend_pos = col23;
                                                                xl_bend_pos = "W";

                                                            }
                                                            if (check17.Contains("{X}") == true)
                                                            {
                                                                col_bend_pos = col24;
                                                                xl_bend_pos = "X";

                                                            }
                                                            if (check17.Contains("{Y}") == true)
                                                            {
                                                                col_bend_pos = col25;
                                                                xl_bend_pos = "Y";

                                                            }
                                                            if (check17.Contains("{Z}") == true)
                                                            {
                                                                col_bend_pos = col26;
                                                                xl_bend_pos = "Z";

                                                            }

                                                        }
                                                        #endregion
                                                    }


                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][18] != DBNull.Value)
                                                    {
                                                        #region BEND horizontal
                                                        string check18 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][18]);
                                                        if (check18.Contains("{F}") == true ||
                                                            check18.Contains("{G}") == true ||
                                                            check18.Contains("{H}") == true ||
                                                            check18.Contains("{I}") == true ||
                                                            check18.Contains("{J}") == true ||
                                                            check18.Contains("{K}") == true ||
                                                            check18.Contains("{L}") == true ||
                                                            check18.Contains("{M}") == true ||
                                                            check18.Contains("{N}") == true ||
                                                            check18.Contains("{O}") == true ||
                                                            check18.Contains("{P}") == true ||
                                                            check18.Contains("{Q}") == true ||
                                                            check18.Contains("{R}") == true ||
                                                            check18.Contains("{S}") == true ||
                                                            check18.Contains("{T}") == true ||
                                                            check18.Contains("{U}") == true ||
                                                            check18.Contains("{V}") == true ||
                                                            check18.Contains("{W}") == true ||
                                                            check18.Contains("{X}") == true ||
                                                            check18.Contains("{Y}") == true ||
                                                            check18.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check18.Contains("{F}") == true)
                                                            {
                                                                col_bend_hor = col6;
                                                                xl_bend_hor = "F";

                                                            }
                                                            if (check18.Contains("{G}") == true)
                                                            {
                                                                col_bend_hor = col7;
                                                                xl_bend_hor = "G";

                                                            }
                                                            if (check18.Contains("{H}") == true)
                                                            {
                                                                col_bend_hor = col8;
                                                                xl_bend_hor = "H";

                                                            }
                                                            if (check18.Contains("{I}") == true)
                                                            {
                                                                col_bend_hor = col9;
                                                                xl_bend_hor = "I";

                                                            }
                                                            if (check18.Contains("{J}") == true)
                                                            {
                                                                col_bend_hor = col10;
                                                                xl_bend_hor = "J";

                                                            }
                                                            if (check18.Contains("{K}") == true)
                                                            {
                                                                col_bend_hor = col11;
                                                                xl_bend_hor = "K";

                                                            }
                                                            if (check18.Contains("{L}") == true)
                                                            {
                                                                col_bend_hor = col12;
                                                                xl_bend_hor = "L";

                                                            }
                                                            if (check18.Contains("{M}") == true)
                                                            {
                                                                col_bend_hor = col13;
                                                                xl_bend_hor = "M";

                                                            }
                                                            if (check18.Contains("{N}") == true)
                                                            {
                                                                col_bend_hor = col14;
                                                                xl_bend_hor = "N";

                                                            }
                                                            if (check18.Contains("{O}") == true)
                                                            {
                                                                col_bend_hor = col15;
                                                                xl_bend_hor = "O";

                                                            }
                                                            if (check18.Contains("{P}") == true)
                                                            {
                                                                col_bend_hor = col16;
                                                                xl_bend_hor = "P";

                                                            }

                                                            if (check18.Contains("{Q}") == true)
                                                            {
                                                                col_bend_hor = col17;
                                                                xl_bend_hor = "Q";
                                                            }
                                                            if (check18.Contains("{R}") == true)
                                                            {
                                                                col_bend_hor = col18;
                                                                xl_bend_hor = "R";
                                                            }
                                                            if (check18.Contains("{S}") == true)
                                                            {
                                                                col_bend_hor = col19;
                                                                xl_bend_hor = "S";
                                                            }
                                                            if (check18.Contains("{T}") == true)
                                                            {
                                                                col_bend_hor = col20;
                                                                xl_bend_hor = "T";
                                                            }
                                                            if (check18.Contains("{U}") == true)
                                                            {
                                                                col_bend_hor = col21;
                                                                xl_bend_hor = "U";
                                                            }
                                                            if (check18.Contains("{V}") == true)
                                                            {
                                                                col_bend_hor = col22;
                                                                xl_bend_hor = "V";
                                                            }
                                                            if (check18.Contains("{W}") == true)
                                                            {
                                                                col_bend_hor = col23;
                                                                xl_bend_hor = "W";
                                                            }
                                                            if (check18.Contains("{X}") == true)
                                                            {
                                                                col_bend_hor = col24;
                                                                xl_bend_hor = "X";
                                                            }
                                                            if (check18.Contains("{Y}") == true)
                                                            {
                                                                col_bend_hor = col25;
                                                                xl_bend_hor = "Y";
                                                            }
                                                            if (check18.Contains("{Z}") == true)
                                                            {
                                                                col_bend_hor = col26;
                                                                xl_bend_hor = "Z";
                                                            }

                                                        }
                                                        #endregion
                                                    }

                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][19] != DBNull.Value)
                                                    {
                                                        #region BEND vertical
                                                        string check19 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][19]);
                                                        if (check19.Contains("{F}") == true ||
                                                            check19.Contains("{G}") == true ||
                                                            check19.Contains("{H}") == true ||
                                                            check19.Contains("{I}") == true ||
                                                            check19.Contains("{J}") == true ||
                                                            check19.Contains("{K}") == true ||
                                                            check19.Contains("{L}") == true ||
                                                            check19.Contains("{M}") == true ||
                                                            check19.Contains("{N}") == true ||
                                                            check19.Contains("{O}") == true ||
                                                            check19.Contains("{P}") == true ||
                                                            check19.Contains("{Q}") == true ||
                                                            check19.Contains("{R}") == true ||
                                                            check19.Contains("{S}") == true ||
                                                            check19.Contains("{T}") == true ||
                                                            check19.Contains("{U}") == true ||
                                                            check19.Contains("{V}") == true ||
                                                            check19.Contains("{W}") == true ||
                                                            check19.Contains("{X}") == true ||
                                                            check19.Contains("{Y}") == true ||
                                                            check19.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check19.Contains("{F}") == true)
                                                            {
                                                                col_bend_ver = col6;
                                                                xl_bend_ver = "F";

                                                            }
                                                            if (check19.Contains("{G}") == true)
                                                            {
                                                                col_bend_ver = col7;
                                                                xl_bend_ver = "G";

                                                            }
                                                            if (check19.Contains("{H}") == true)
                                                            {
                                                                col_bend_ver = col8;
                                                                xl_bend_ver = "H";

                                                            }
                                                            if (check19.Contains("{I}") == true)
                                                            {
                                                                col_bend_ver = col9;
                                                                xl_bend_ver = "I";

                                                            }
                                                            if (check19.Contains("{J}") == true)
                                                            {
                                                                col_bend_ver = col10;
                                                                xl_bend_ver = "J";

                                                            }
                                                            if (check19.Contains("{K}") == true)
                                                            {
                                                                col_bend_ver = col11;
                                                                xl_bend_ver = "K";

                                                            }
                                                            if (check19.Contains("{L}") == true)
                                                            {
                                                                col_bend_ver = col12;
                                                                xl_bend_ver = "L";

                                                            }
                                                            if (check19.Contains("{M}") == true)
                                                            {
                                                                col_bend_ver = col13;
                                                                xl_bend_ver = "M";

                                                            }
                                                            if (check19.Contains("{N}") == true)
                                                            {
                                                                col_bend_ver = col14;
                                                                xl_bend_ver = "N";

                                                            }
                                                            if (check19.Contains("{O}") == true)
                                                            {
                                                                col_bend_ver = col15;
                                                                xl_bend_ver = "O";

                                                            }
                                                            if (check19.Contains("{P}") == true)
                                                            {
                                                                col_bend_ver = col16;
                                                                xl_bend_ver = "P";

                                                            }

                                                            if (check19.Contains("{Q}") == true)
                                                            {
                                                                col_bend_ver = col17;
                                                                xl_bend_ver = "Q";
                                                            }
                                                            if (check19.Contains("{R}") == true)
                                                            {
                                                                col_bend_ver = col18;
                                                                xl_bend_ver = "R";
                                                            }
                                                            if (check19.Contains("{S}") == true)
                                                            {
                                                                col_bend_ver = col19;
                                                                xl_bend_ver = "S";
                                                            }
                                                            if (check19.Contains("{T}") == true)
                                                            {
                                                                col_bend_ver = col20;
                                                                xl_bend_ver = "T";
                                                            }
                                                            if (check19.Contains("{U}") == true)
                                                            {
                                                                col_bend_ver = col21;
                                                                xl_bend_ver = "U";
                                                            }
                                                            if (check19.Contains("{V}") == true)
                                                            {
                                                                col_bend_ver = col22;
                                                                xl_bend_ver = "V";
                                                            }
                                                            if (check19.Contains("{W}") == true)
                                                            {
                                                                col_bend_ver = col23;
                                                                xl_bend_ver = "W";
                                                            }
                                                            if (check19.Contains("{X}") == true)
                                                            {
                                                                col_bend_ver = col24;
                                                                xl_bend_ver = "X";
                                                            }
                                                            if (check19.Contains("{Y}") == true)
                                                            {
                                                                col_bend_ver = col25;
                                                                xl_bend_ver = "Y";
                                                            }
                                                            if (check19.Contains("{Z}") == true)
                                                            {
                                                                col_bend_ver = col26;
                                                                xl_bend_ver = "Z";
                                                            }

                                                        }
                                                        #endregion
                                                    }

                                                    #endregion                                                  
                                                }
                                               
                                                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "ELBOW")
                                                {
                                                    #region ELBOW
                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][15] != DBNull.Value)
                                                    {
                                                        #region ELBOW TYPE
                                                        string check15 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][15]);
                                                        if (check15.Contains("{F}") == true ||
                                                            check15.Contains("{G}") == true ||
                                                            check15.Contains("{H}") == true ||
                                                            check15.Contains("{I}") == true ||
                                                            check15.Contains("{J}") == true ||
                                                            check15.Contains("{K}") == true ||
                                                            check15.Contains("{L}") == true ||
                                                            check15.Contains("{M}") == true ||
                                                            check15.Contains("{N}") == true ||
                                                            check15.Contains("{O}") == true ||
                                                            check15.Contains("{P}") == true ||
                                                            check15.Contains("{Q}") == true ||
                                                            check15.Contains("{R}") == true ||
                                                            check15.Contains("{S}") == true ||
                                                            check15.Contains("{T}") == true ||
                                                            check15.Contains("{U}") == true ||
                                                            check15.Contains("{V}") == true ||
                                                            check15.Contains("{W}") == true ||
                                                            check15.Contains("{X}") == true ||
                                                            check15.Contains("{Y}") == true ||
                                                            check15.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check15.Contains("{F}") == true)
                                                            {
                                                                col_elbow_type = col6;
                                                                xl_elbow_type = "F";

                                                            }
                                                            if (check15.Contains("{G}") == true)
                                                            {
                                                                col_elbow_type = col7;
                                                                xl_elbow_type = "G";

                                                            }
                                                            if (check15.Contains("{H}") == true)
                                                            {
                                                                col_elbow_type = col8;
                                                                xl_elbow_type = "H";

                                                            }
                                                            if (check15.Contains("{I}") == true)
                                                            {
                                                                col_elbow_type = col9;
                                                                xl_elbow_type = "I";

                                                            }
                                                            if (check15.Contains("{J}") == true)
                                                            {
                                                                col_elbow_type = col10;
                                                                xl_elbow_type = "J";

                                                            }
                                                            if (check15.Contains("{K}") == true)
                                                            {
                                                                col_elbow_type = col11;
                                                                xl_elbow_type = "K";

                                                            }
                                                            if (check15.Contains("{L}") == true)
                                                            {
                                                                col_elbow_type = col12;
                                                                xl_elbow_type = "L";

                                                            }
                                                            if (check15.Contains("{M}") == true)
                                                            {
                                                                col_elbow_type = col13;
                                                                xl_elbow_type = "M";

                                                            }
                                                            if (check15.Contains("{N}") == true)
                                                            {
                                                                col_elbow_type = col14;
                                                                xl_elbow_type = "N";

                                                            }
                                                            if (check15.Contains("{O}") == true)
                                                            {
                                                                col_elbow_type = col15;
                                                                xl_elbow_type = "O";

                                                            }
                                                            if (check15.Contains("{P}") == true)
                                                            {
                                                                col_elbow_type = col16;
                                                                xl_elbow_type = "P";

                                                            }
                                                            if (check15.Contains("{Q}") == true)
                                                            {
                                                                col_elbow_type = col17;
                                                                xl_elbow_type = "Q";

                                                            }
                                                            if (check15.Contains("{R}") == true)
                                                            {
                                                                col_elbow_type = col18;
                                                                xl_elbow_type = "R";

                                                            }
                                                            if (check15.Contains("{S}") == true)
                                                            {
                                                                col_elbow_type = col19;
                                                                xl_elbow_type = "S";

                                                            }
                                                            if (check15.Contains("{T}") == true)
                                                            {
                                                                col_elbow_type = col20;
                                                                xl_elbow_type = "T";

                                                            }
                                                            if (check15.Contains("{U}") == true)
                                                            {
                                                                col_elbow_type = col21;
                                                                xl_elbow_type = "U";

                                                            }
                                                            if (check15.Contains("{V}") == true)
                                                            {
                                                                col_elbow_type = col22;
                                                                xl_elbow_type = "V";

                                                            }
                                                            if (check15.Contains("{W}") == true)
                                                            {
                                                                col_elbow_type = col23;
                                                                xl_elbow_type = "W";

                                                            }
                                                            if (check15.Contains("{X}") == true)
                                                            {
                                                                col_elbow_type = col24;
                                                                xl_elbow_type = "X";

                                                            }
                                                            if (check15.Contains("{Y}") == true)
                                                            {
                                                                col_elbow_type = col25;
                                                                xl_elbow_type = "Y";

                                                            }
                                                            if (check15.Contains("{Z}") == true)
                                                            {
                                                                col_elbow_type = col26;
                                                                xl_elbow_type = "Z";

                                                            }



                                                        }
                                                        #endregion
                                                    }

                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][16] != DBNull.Value)
                                                    {
                                                        #region ELBOW DEFLECTION
                                                        string check16 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][16]);
                                                        if (check16.Contains("{F}") == true ||
                                                            check16.Contains("{G}") == true ||
                                                            check16.Contains("{H}") == true ||
                                                            check16.Contains("{I}") == true ||
                                                            check16.Contains("{J}") == true ||
                                                            check16.Contains("{K}") == true ||
                                                            check16.Contains("{L}") == true ||
                                                            check16.Contains("{M}") == true ||
                                                            check16.Contains("{N}") == true ||
                                                            check16.Contains("{O}") == true ||
                                                            check16.Contains("{P}") == true ||
                                                            check16.Contains("{Q}") == true ||
                                                            check16.Contains("{R}") == true ||
                                                            check16.Contains("{S}") == true ||
                                                            check16.Contains("{T}") == true ||
                                                            check16.Contains("{U}") == true ||
                                                            check16.Contains("{V}") == true ||
                                                            check16.Contains("{W}") == true ||
                                                            check16.Contains("{X}") == true ||
                                                            check16.Contains("{Y}") == true ||
                                                            check16.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check16.Contains("{F}") == true)
                                                            {
                                                                col_elbow_defl = col6;
                                                                xl_elbow_defl = "F";

                                                            }
                                                            if (check16.Contains("{G}") == true)
                                                            {
                                                                col_elbow_defl = col7;
                                                                xl_elbow_defl = "G";

                                                            }
                                                            if (check16.Contains("{H}") == true)
                                                            {
                                                                col_elbow_defl = col8;
                                                                xl_elbow_defl = "H";

                                                            }
                                                            if (check16.Contains("{I}") == true)
                                                            {
                                                                col_elbow_defl = col9;
                                                                xl_elbow_defl = "I";

                                                            }
                                                            if (check16.Contains("{J}") == true)
                                                            {
                                                                col_elbow_defl = col10;
                                                                xl_elbow_defl = "J";

                                                            }
                                                            if (check16.Contains("{K}") == true)
                                                            {
                                                                col_elbow_defl = col11;
                                                                xl_elbow_defl = "K";

                                                            }
                                                            if (check16.Contains("{L}") == true)
                                                            {
                                                                col_elbow_defl = col12;
                                                                xl_elbow_defl = "L";

                                                            }
                                                            if (check16.Contains("{M}") == true)
                                                            {
                                                                col_elbow_defl = col13;
                                                                xl_elbow_defl = "M";

                                                            }
                                                            if (check16.Contains("{N}") == true)
                                                            {
                                                                col_elbow_defl = col14;
                                                                xl_elbow_defl = "N";

                                                            }
                                                            if (check16.Contains("{O}") == true)
                                                            {
                                                                col_elbow_defl = col15;
                                                                xl_elbow_defl = "O";

                                                            }
                                                            if (check16.Contains("{P}") == true)
                                                            {
                                                                col_elbow_defl = col16;
                                                                xl_elbow_defl = "P";

                                                            }

                                                            if (check16.Contains("{Q}") == true)
                                                            {
                                                                col_elbow_defl = col17;
                                                                xl_elbow_defl = "Q";

                                                            }
                                                            if (check16.Contains("{R}") == true)
                                                            {
                                                                col_elbow_defl = col18;
                                                                xl_elbow_defl = "R";

                                                            }
                                                            if (check16.Contains("{S}") == true)
                                                            {
                                                                col_elbow_defl = col19;
                                                                xl_elbow_defl = "S";

                                                            }
                                                            if (check16.Contains("{T}") == true)
                                                            {
                                                                col_elbow_defl = col20;
                                                                xl_elbow_defl = "T";

                                                            }
                                                            if (check16.Contains("{U}") == true)
                                                            {
                                                                col_elbow_defl = col21;
                                                                xl_elbow_defl = "U";

                                                            }
                                                            if (check16.Contains("{V}") == true)
                                                            {
                                                                col_elbow_defl = col22;
                                                                xl_elbow_defl = "V";

                                                            }
                                                            if (check16.Contains("{W}") == true)
                                                            {
                                                                col_elbow_defl = col23;
                                                                xl_elbow_defl = "W";

                                                            }
                                                            if (check16.Contains("{X}") == true)
                                                            {
                                                                col_elbow_defl = col24;
                                                                xl_elbow_defl = "X";

                                                            }
                                                            if (check16.Contains("{Y}") == true)
                                                            {
                                                                col_elbow_defl = col25;
                                                                xl_elbow_defl = "Y";

                                                            }
                                                            if (check16.Contains("{Z}") == true)
                                                            {
                                                                col_elbow_defl = col26;
                                                                xl_elbow_defl = "Z";

                                                            }


                                                        }
                                                        #endregion
                                                    }




                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][17] != DBNull.Value)
                                                    {
                                                        #region ELBOW POSITION
                                                        string check17 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][17]);
                                                        if (check17.Contains("{F}") == true ||
                                                            check17.Contains("{G}") == true ||
                                                            check17.Contains("{H}") == true ||
                                                            check17.Contains("{I}") == true ||
                                                            check17.Contains("{J}") == true ||
                                                            check17.Contains("{K}") == true ||
                                                            check17.Contains("{L}") == true ||
                                                            check17.Contains("{M}") == true ||
                                                            check17.Contains("{N}") == true ||
                                                            check17.Contains("{O}") == true ||
                                                            check17.Contains("{P}") == true ||
                                                            check17.Contains("{Q}") == true ||
                                                            check17.Contains("{R}") == true ||
                                                            check17.Contains("{S}") == true ||
                                                            check17.Contains("{T}") == true ||
                                                            check17.Contains("{U}") == true ||
                                                            check17.Contains("{V}") == true ||
                                                            check17.Contains("{W}") == true ||
                                                            check17.Contains("{X}") == true ||
                                                            check17.Contains("{Y}") == true ||
                                                            check17.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check17.Contains("{F}") == true)
                                                            {
                                                                col_elbow_pos = col6;
                                                                xl_elbow_pos = "F";

                                                            }
                                                            if (check17.Contains("{G}") == true)
                                                            {
                                                                col_elbow_pos = col7;
                                                                xl_elbow_pos = "G";

                                                            }
                                                            if (check17.Contains("{H}") == true)
                                                            {
                                                                col_elbow_pos = col8;
                                                                xl_elbow_pos = "H";

                                                            }
                                                            if (check17.Contains("{I}") == true)
                                                            {
                                                                col_elbow_pos = col9;
                                                                xl_elbow_pos = "I";

                                                            }
                                                            if (check17.Contains("{J}") == true)
                                                            {
                                                                col_elbow_pos = col10;
                                                                xl_elbow_pos = "J";

                                                            }
                                                            if (check17.Contains("{K}") == true)
                                                            {
                                                                col_elbow_pos = col11;
                                                                xl_elbow_pos = "K";

                                                            }
                                                            if (check17.Contains("{L}") == true)
                                                            {
                                                                col_elbow_pos = col12;
                                                                xl_elbow_pos = "L";

                                                            }
                                                            if (check17.Contains("{M}") == true)
                                                            {
                                                                col_elbow_pos = col13;

                                                            }
                                                            if (check17.Contains("{N}") == true)
                                                            {
                                                                col_elbow_pos = col14;
                                                                xl_elbow_pos = "N";

                                                            }
                                                            if (check17.Contains("{O}") == true)
                                                            {
                                                                col_elbow_pos = col15;
                                                                xl_elbow_pos = "O";

                                                            }
                                                            if (check17.Contains("{P}") == true)
                                                            {
                                                                col_elbow_pos = col16;
                                                                xl_elbow_pos = "P";

                                                            }
                                                            if (check17.Contains("{Q}") == true)
                                                            {
                                                                col_elbow_pos = col17;
                                                                xl_elbow_pos = "Q";

                                                            }
                                                            if (check17.Contains("{R}") == true)
                                                            {
                                                                col_elbow_pos = col18;
                                                                xl_elbow_pos = "R";

                                                            }
                                                            if (check17.Contains("{S}") == true)
                                                            {
                                                                col_elbow_pos = col19;
                                                                xl_elbow_pos = "S";

                                                            }
                                                            if (check17.Contains("{T}") == true)
                                                            {
                                                                col_elbow_pos = col20;
                                                                xl_elbow_pos = "T";

                                                            }
                                                            if (check17.Contains("{U}") == true)
                                                            {
                                                                col_elbow_pos = col21;
                                                                xl_elbow_pos = "U";

                                                            }
                                                            if (check17.Contains("{V}") == true)
                                                            {
                                                                col_elbow_pos = col22;
                                                                xl_elbow_pos = "V";

                                                            }
                                                            if (check17.Contains("{W}") == true)
                                                            {
                                                                col_elbow_pos = col23;
                                                                xl_elbow_pos = "W";

                                                            }
                                                            if (check17.Contains("{X}") == true)
                                                            {
                                                                col_elbow_pos = col24;
                                                                xl_elbow_pos = "X";

                                                            }
                                                            if (check17.Contains("{Y}") == true)
                                                            {
                                                                col_elbow_pos = col25;
                                                                xl_elbow_pos = "Y";

                                                            }
                                                            if (check17.Contains("{Z}") == true)
                                                            {
                                                                col_elbow_pos = col26;
                                                                xl_elbow_pos = "Z";

                                                            }

                                                        }
                                                        #endregion
                                                    }


                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][18] != DBNull.Value)
                                                    {
                                                        #region ELBOW horizontal
                                                        string check18 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][18]);
                                                        if (check18.Contains("{F}") == true ||
                                                            check18.Contains("{G}") == true ||
                                                            check18.Contains("{H}") == true ||
                                                            check18.Contains("{I}") == true ||
                                                            check18.Contains("{J}") == true ||
                                                            check18.Contains("{K}") == true ||
                                                            check18.Contains("{L}") == true ||
                                                            check18.Contains("{M}") == true ||
                                                            check18.Contains("{N}") == true ||
                                                            check18.Contains("{O}") == true ||
                                                            check18.Contains("{P}") == true ||
                                                            check18.Contains("{Q}") == true ||
                                                            check18.Contains("{R}") == true ||
                                                            check18.Contains("{S}") == true ||
                                                            check18.Contains("{T}") == true ||
                                                            check18.Contains("{U}") == true ||
                                                            check18.Contains("{V}") == true ||
                                                            check18.Contains("{W}") == true ||
                                                            check18.Contains("{X}") == true ||
                                                            check18.Contains("{Y}") == true ||
                                                            check18.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check18.Contains("{F}") == true)
                                                            {
                                                                col_elbow_hor = col6;
                                                                xl_elbow_hor = "F";

                                                            }
                                                            if (check18.Contains("{G}") == true)
                                                            {
                                                                col_elbow_hor = col7;
                                                                xl_elbow_hor = "G";

                                                            }
                                                            if (check18.Contains("{H}") == true)
                                                            {
                                                                col_elbow_hor = col8;
                                                                xl_elbow_hor = "H";

                                                            }
                                                            if (check18.Contains("{I}") == true)
                                                            {
                                                                col_elbow_hor = col9;
                                                                xl_elbow_hor = "I";

                                                            }
                                                            if (check18.Contains("{J}") == true)
                                                            {
                                                                col_elbow_hor = col10;
                                                                xl_elbow_hor = "J";

                                                            }
                                                            if (check18.Contains("{K}") == true)
                                                            {
                                                                col_elbow_hor = col11;
                                                                xl_elbow_hor = "K";

                                                            }
                                                            if (check18.Contains("{L}") == true)
                                                            {
                                                                col_elbow_hor = col12;
                                                                xl_elbow_hor = "L";

                                                            }
                                                            if (check18.Contains("{M}") == true)
                                                            {
                                                                col_elbow_hor = col13;
                                                                xl_elbow_hor = "M";

                                                            }
                                                            if (check18.Contains("{N}") == true)
                                                            {
                                                                col_elbow_hor = col14;
                                                                xl_elbow_hor = "N";

                                                            }
                                                            if (check18.Contains("{O}") == true)
                                                            {
                                                                col_elbow_hor = col15;
                                                                xl_elbow_hor = "O";

                                                            }
                                                            if (check18.Contains("{P}") == true)
                                                            {
                                                                col_elbow_hor = col16;
                                                                xl_elbow_hor = "P";

                                                            }

                                                            if (check18.Contains("{Q}") == true)
                                                            {
                                                                col_elbow_hor = col17;
                                                                xl_elbow_hor = "Q";
                                                            }
                                                            if (check18.Contains("{R}") == true)
                                                            {
                                                                col_elbow_hor = col18;
                                                                xl_elbow_hor = "R";
                                                            }
                                                            if (check18.Contains("{S}") == true)
                                                            {
                                                                col_elbow_hor = col19;
                                                                xl_elbow_hor = "S";
                                                            }
                                                            if (check18.Contains("{T}") == true)
                                                            {
                                                                col_elbow_hor = col20;
                                                                xl_elbow_hor = "T";
                                                            }
                                                            if (check18.Contains("{U}") == true)
                                                            {
                                                                col_elbow_hor = col21;
                                                                xl_elbow_hor = "U";
                                                            }
                                                            if (check18.Contains("{V}") == true)
                                                            {
                                                                col_elbow_hor = col22;
                                                                xl_elbow_hor = "V";
                                                            }
                                                            if (check18.Contains("{W}") == true)
                                                            {
                                                                col_elbow_hor = col23;
                                                                xl_elbow_hor = "W";
                                                            }
                                                            if (check18.Contains("{X}") == true)
                                                            {
                                                                col_elbow_hor = col24;
                                                                xl_elbow_hor = "X";
                                                            }
                                                            if (check18.Contains("{Y}") == true)
                                                            {
                                                                col_elbow_hor = col25;
                                                                xl_elbow_hor = "Y";
                                                            }
                                                            if (check18.Contains("{Z}") == true)
                                                            {
                                                                col_elbow_hor = col26;
                                                                xl_elbow_hor = "Z";
                                                            }

                                                        }
                                                        #endregion
                                                    }

                                                    if (Wgen_main_form.dt_feature_codes.Rows[j][19] != DBNull.Value)
                                                    {
                                                        #region ELBOW vertical
                                                        string check19 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][19]);
                                                        if (check19.Contains("{F}") == true ||
                                                            check19.Contains("{G}") == true ||
                                                            check19.Contains("{H}") == true ||
                                                            check19.Contains("{I}") == true ||
                                                            check19.Contains("{J}") == true ||
                                                            check19.Contains("{K}") == true ||
                                                            check19.Contains("{L}") == true ||
                                                            check19.Contains("{M}") == true ||
                                                            check19.Contains("{N}") == true ||
                                                            check19.Contains("{O}") == true ||
                                                            check19.Contains("{P}") == true ||
                                                            check19.Contains("{Q}") == true ||
                                                            check19.Contains("{R}") == true ||
                                                            check19.Contains("{S}") == true ||
                                                            check19.Contains("{T}") == true ||
                                                            check19.Contains("{U}") == true ||
                                                            check19.Contains("{V}") == true ||
                                                            check19.Contains("{W}") == true ||
                                                            check19.Contains("{X}") == true ||
                                                            check19.Contains("{Y}") == true ||
                                                            check19.Contains("{Z}") == true
                                                            )
                                                        {
                                                            if (check19.Contains("{F}") == true)
                                                            {
                                                                col_elbow_ver = col6;
                                                                xl_elbow_ver = "F";

                                                            }
                                                            if (check19.Contains("{G}") == true)
                                                            {
                                                                col_elbow_ver = col7;
                                                                xl_elbow_ver = "G";

                                                            }
                                                            if (check19.Contains("{H}") == true)
                                                            {
                                                                col_elbow_ver = col8;
                                                                xl_elbow_ver = "H";

                                                            }
                                                            if (check19.Contains("{I}") == true)
                                                            {
                                                                col_elbow_ver = col9;
                                                                xl_elbow_ver = "I";

                                                            }
                                                            if (check19.Contains("{J}") == true)
                                                            {
                                                                col_elbow_ver = col10;
                                                                xl_elbow_ver = "J";

                                                            }
                                                            if (check19.Contains("{K}") == true)
                                                            {
                                                                col_elbow_ver = col11;
                                                                xl_elbow_ver = "K";

                                                            }
                                                            if (check19.Contains("{L}") == true)
                                                            {
                                                                col_elbow_ver = col12;
                                                                xl_elbow_ver = "L";

                                                            }
                                                            if (check19.Contains("{M}") == true)
                                                            {
                                                                col_elbow_ver = col13;
                                                                xl_elbow_ver = "M";

                                                            }
                                                            if (check19.Contains("{N}") == true)
                                                            {
                                                                col_elbow_ver = col14;
                                                                xl_elbow_ver = "N";

                                                            }
                                                            if (check19.Contains("{O}") == true)
                                                            {
                                                                col_elbow_ver = col15;
                                                                xl_elbow_ver = "O";

                                                            }
                                                            if (check19.Contains("{P}") == true)
                                                            {
                                                                col_elbow_ver = col16;
                                                                xl_elbow_ver = "P";

                                                            }

                                                            if (check19.Contains("{Q}") == true)
                                                            {
                                                                col_elbow_ver = col17;
                                                                xl_elbow_ver = "Q";
                                                            }
                                                            if (check19.Contains("{R}") == true)
                                                            {
                                                                col_elbow_ver = col18;
                                                                xl_elbow_ver = "R";
                                                            }
                                                            if (check19.Contains("{S}") == true)
                                                            {
                                                                col_elbow_ver = col19;
                                                                xl_elbow_ver = "S";
                                                            }
                                                            if (check19.Contains("{T}") == true)
                                                            {
                                                                col_elbow_ver = col20;
                                                                xl_elbow_ver = "T";
                                                            }
                                                            if (check19.Contains("{U}") == true)
                                                            {
                                                                col_elbow_ver = col21;
                                                                xl_elbow_ver = "U";
                                                            }
                                                            if (check19.Contains("{V}") == true)
                                                            {
                                                                col_elbow_ver = col22;
                                                                xl_elbow_ver = "V";
                                                            }
                                                            if (check19.Contains("{W}") == true)
                                                            {
                                                                col_elbow_ver = col23;
                                                                xl_elbow_ver = "W";
                                                            }
                                                            if (check19.Contains("{X}") == true)
                                                            {
                                                                col_elbow_ver = col24;
                                                                xl_elbow_ver = "X";
                                                            }
                                                            if (check19.Contains("{Y}") == true)
                                                            {
                                                                col_elbow_ver = col25;
                                                                xl_elbow_ver = "Y";
                                                            }
                                                            if (check19.Contains("{Z}") == true)
                                                            {
                                                                col_elbow_ver = col26;
                                                                xl_elbow_ver = "Z";
                                                            }

                                                        }
                                                        #endregion
                                                    }

                                                    #endregion                                                  
                                                }



                                            }
                                        }
                                    }
                                }



                                var pt_duplicates = Wgen_main_form.dt_all_points.AsEnumerable().GroupBy(i => new { Ptno = i.Field<string>(col1) }).Where(g => g.Count() > 1).Select(g => new { g.Key.Ptno }).ToList();
                                var xray_duplicates = Wgen_main_form.dt_all_points.AsEnumerable().GroupBy(i => new { desc1 = i.Field<string>(col_xray), feat1 = i.Field<string>(col5) }).
                                    Where(g => g.Count() > 1).Select(g => new { g.Key.desc1, g.Key.feat1 }).ToList();


                                System.Data.DataTable dt2 = new System.Data.DataTable();
                                dt2.Columns.Add(col1, typeof(string));

                                System.Data.DataTable dt3 = new System.Data.DataTable();
                                dt3.Columns.Add(col_xray, typeof(string));


                                if (Wgen_main_form.dt_ground_tally != null) Wgen_main_form.dt_ground_tally.TableName = "t1";
                                DataSet dataset1 = new DataSet();
                                dataset1.Tables.Add(Wgen_main_form.dt_all_points);

                                if (pt_duplicates.Count > 0 || xray_duplicates.Count > 0)
                                {
                                    if (pt_duplicates.Count > 0)
                                    {
                                        for (int i = 0; i < pt_duplicates.Count; ++i)
                                        {
                                            if (pt_duplicates[i].Ptno != null)
                                            {
                                                string duplicat_val1 = Convert.ToString(pt_duplicates[i].Ptno);
                                                dt2.Rows.Add();
                                                dt2.Rows[dt2.Rows.Count - 1][0] = duplicat_val1;

                                            }
                                        }
                                    }

                                    if (xray_duplicates.Count > 0)
                                    {
                                        for (int i = 0; i < xray_duplicates.Count; ++i)
                                        {
                                            if (xray_duplicates[i].desc1 != null && xray_duplicates[i].feat1 != null)
                                            {
                                                string duplicat_val1 = Convert.ToString(xray_duplicates[i].desc1);
                                                string feature1 = Convert.ToString(xray_duplicates[i].feat1);
                                                if (feature1.ToUpper() == "WELD" || feature1.ToUpper() == "WLD")
                                                {
                                                    dt3.Rows.Add();
                                                    dt3.Rows[dt3.Rows.Count - 1][0] = duplicat_val1;
                                                }

                                            }
                                        }


                                    }


                                    if (pt_duplicates.Count > 0) dataset1.Tables.Add(dt2);
                                    if (xray_duplicates.Count > 0) dataset1.Tables.Add(dt3);

                                    DataRelation relation1 = null;
                                    if (pt_duplicates.Count > 0) relation1 = new DataRelation("xxx", Wgen_main_form.dt_all_points.Columns[col1], dt2.Columns[col1], false);

                                    DataRelation relation2 = null;
                                    if (xray_duplicates.Count > 0) relation2 = new DataRelation("xx1x", Wgen_main_form.dt_all_points.Columns[col_xray], dt3.Columns[col_xray], false);

                                    if (pt_duplicates.Count > 0) dataset1.Relations.Add(relation1);
                                    if (xray_duplicates.Count > 0) dataset1.Relations.Add(relation2);

                                    nr_duplicates = dt2.Rows.Count;

                                    for (int i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                                    {
                                        if (pt_duplicates.Count > 0)
                                        {
                                            if (Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1).Length > 0)
                                            {
                                                for (int j = 0; j < Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1).Length; ++j)
                                                {
                                                    string val2 = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1)[j][col1].ToString();
                                                    string val1 = Wgen_main_form.dt_all_points.Rows[i][col1].ToString();
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = val2;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_1.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Duplicate Point Number";

                                                    string x = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                    }
                                                    string y = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                    }

                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;

                                                }
                                            }
                                        }

                                        if (xray_duplicates.Count > 0)
                                        {
                                            if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value &&
                                                (Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WELD" || Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WLD") &&
                                                Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation2).Length > 0)
                                            {
                                                if (Wgen_main_form.lista_feature_code_exception == null ||
                                                    Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                                    Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                                {
                                                    string val9 = Wgen_main_form.dt_all_points.Rows[i][col_xray].ToString();
                                                    string val1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col1]);

                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = val1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = val9;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_xray + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "Duplicate Xray Number";
                                                    string x = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                    }
                                                    string y = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                    }

                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;

                                                    ++nr_xray_duplicates;
                                                }


                                            }
                                        }
                                    }


                                    if (pt_duplicates.Count > 0)
                                    {
                                        dataset1.Relations.Remove(relation1);
                                        dataset1.Tables.Remove(dt2);
                                        dt2 = null;
                                    }

                                    if (xray_duplicates.Count > 0)
                                    {
                                        dataset1.Relations.Remove(relation2);
                                        dataset1.Tables.Remove(dt3);
                                        dt3 = null;
                                    }

                                }


                                if (Wgen_main_form.dt_ground_tally != null && Wgen_main_form.dt_ground_tally.Rows.Count > 0)
                                {
                                    dataset1.Tables.Add(Wgen_main_form.dt_ground_tally);
                                    DataRelation relation1 = new DataRelation("xxxY", Wgen_main_form.dt_all_points.Columns[col_mm_back], Wgen_main_form.dt_ground_tally.Columns[colGT1], false);
                                    dataset1.Relations.Add(relation1);
                                    DataRelation relation2 = new DataRelation("xxxZ", Wgen_main_form.dt_all_points.Columns[col_mm_ahead], Wgen_main_form.dt_ground_tally.Columns[colGT1], false);
                                    dataset1.Relations.Add(relation2);
                                    for (int i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                                    {
                                        string mmid1 = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col_mm_back] != DBNull.Value)
                                        {
                                            mmid1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_mm_back]);
                                        }

                                        string mmid2 = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col_mm_ahead] != DBNull.Value)
                                        {
                                            mmid2 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_mm_ahead]);
                                        }

                                        string pt1 = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col1] != DBNull.Value)
                                        {
                                            pt1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col1]);
                                        }

                                        if (Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1).Length == 0 && mmid1.ToUpper() != "FAB")
                                        {


                                            if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WELD"
                                                || Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WLD")
                                            {
                                                if (Wgen_main_form.lista_feature_code_exception == null ||
                                                    Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                                    Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = mmid1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_mm_back + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "MMid Back not found in Pipe Tally";
                                                    string x = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                    }
                                                    string y = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                    }

                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                    ++nr_mmid_not_found;
                                                }
                                            }
                                        }
                                        if (Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation2).Length == 0 && mmid2.ToUpper() != "FAB")
                                        {
                                            if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WELD" ||
                                                Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "WLD")
                                            {
                                                if (Wgen_main_form.lista_feature_code_exception == null ||
                                                    Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                                    Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Value"] = mmid2;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_mm_ahead + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "MMid Ahead not found in Pipe Tally";
                                                    string x = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                    }
                                                    string y = "";
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                    }

                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                    ++nr_mmid_not_found;
                                                }
                                            }
                                        }
                                    }
                                    dataset1.Relations.Remove(relation1);
                                    dataset1.Relations.Remove(relation2);
                                    dataset1.Tables.Remove(Wgen_main_form.dt_ground_tally);
                                }

                                dataset1.Tables.Remove(Wgen_main_form.dt_all_points);

                                for (int i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                                {
                                    string pt1 = "";
                                    if (Wgen_main_form.dt_all_points.Rows[i][col1] != DBNull.Value)
                                    {
                                        pt1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col1]);
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i][col1] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_1.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Point Number Specified";
                                        string x = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                        {
                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                        }
                                        string y = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                        {
                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                        }

                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                        ++nr_null_values;
                                    }
                                    if (Wgen_main_form.dt_all_points.Rows[i][col2] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_2.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Northing Specified";

                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i][col3] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_3.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Easting Specified";
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i][col4] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_4.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Elevation Specified";
                                        string x = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                        {
                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                        }
                                        string y = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                        {
                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                        }

                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                        ++nr_null_values;
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i][col5] == DBNull.Value)
                                    {
                                        dt_errors.Rows.Add();
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = textBox_5.Text + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Feature Code Specified";
                                        string x = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                        {
                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                        }
                                        string y = "";
                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                        {
                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                        }

                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                        ++nr_null_values;
                                    }


                                    string feature_code = "X";
                                    if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value)
                                    {
                                        feature_code = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]);
                                    }




                                    #region bend checks
                                    if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "BEND")
                                    {
                                        if (Wgen_main_form.lista_feature_code_exception == null ||
                                            Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                            Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                        {
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_type] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_type + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Bend Type Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_defl] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_defl + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Deflection Type Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_pos] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_pos + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Position Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_type] != DBNull.Value && Wgen_main_form.dt_all_points.Rows[i][col_bend_defl] != DBNull.Value)
                                            {
                                                string bend_type1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_type]).ToUpper();
                                                string defl_type = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_defl]).ToUpper();

                                                if (defl_type == "RIGHT" || defl_type == "LEFT")
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_bend_hor] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Horizontal Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                                else if (defl_type == "SAG" || defl_type == "OVERBEND")
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_bend_ver] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_ver + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Vertical Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                                else
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_bend_hor] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Horizontal Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_bend_ver] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_bend_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Vertical Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion


                                    #region elbow checks
                                    if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value && Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5]).ToUpper() == "ELBOW")
                                    {
                                        if (Wgen_main_form.lista_feature_code_exception == null ||
                                            Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                            Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                        {
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_type] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_type + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No elbow Type Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_defl] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_defl + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Deflection Type Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_pos] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_pos + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Position Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_type] != DBNull.Value && Wgen_main_form.dt_all_points.Rows[i][col_elbow_defl] != DBNull.Value)
                                            {
                                                string elbow_type1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_type]).ToUpper();
                                                string defl_type = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_defl]).ToUpper();

                                                if (defl_type == "RIGHT" || defl_type == "LEFT")
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_hor] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Horizontal Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                                else if (defl_type == "SAG" || defl_type == "OVERBEND")
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_ver] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_ver + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Vertical Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                                else
                                                {
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_hor] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Horizontal Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                    if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_ver] == DBNull.Value)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_elbow_hor + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No Vertical Deflection Specified";
                                                        string x = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                        }

                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                        dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                        ++nr_null_values;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion



                                    #region checks inside all points
                                    if (Wgen_main_form.dt_all_points.Rows[i][col5] != DBNull.Value)
                                    {
                                        if (Wgen_main_form.lista_feature_code_exception == null ||
                                            Wgen_main_form.lista_feature_code_exception.Count == 0 ||
                                            Wgen_main_form.lista_feature_code_exception.Contains(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col5])) == false)
                                        {
                                            int col_check1 = 0;
                                            int col_check2 = 0;
                                            int col_check3 = 0;
                                            int col_check4 = 0;
                                            int col_check5 = 0;
                                            int col_check6 = 0;
                                            int col_check7 = 0;
                                            int col_check8 = 0;
                                            int col_check9 = 0;
                                            int col_check10 = 0;

                                            string xl_col1 = "A";
                                            string xl_col2 = "A";
                                            string xl_col3 = "A";
                                            string xl_col4 = "A";
                                            string xl_col5 = "A";
                                            string xl_col6 = "A";
                                            string xl_col7 = "A";
                                            string xl_col8 = "A";
                                            string xl_col9 = "A";
                                            string xl_col10 = "A";

                                            bool number_check1 = false;
                                            bool number_check2 = false;
                                            bool number_check3 = false;
                                            bool number_check4 = false;
                                            bool number_check5 = false;
                                            bool number_check6 = false;
                                            bool number_check7 = false;
                                            bool number_check8 = false;
                                            bool number_check9 = false;
                                            bool number_check10 = false;

                                            string ftype1 = "X";
                                            string ftype2 = "X";
                                            string ftype3 = "X";
                                            string ftype4 = "X";
                                            string ftype5 = "X";
                                            string ftype6 = "X";
                                            string ftype7 = "X";
                                            string ftype8 = "X";
                                            string ftype9 = "X";
                                            string ftype10 = "X";

                                            string fc = "y";

                                            for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                                            {
                                                string client_name = "xxx";
                                                if (Wgen_main_form.dt_feature_codes.Rows[j][0] != DBNull.Value)
                                                {
                                                    client_name = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]);
                                                    if (client_name == Wgen_main_form.client_name)
                                                    {
                                                        if (Wgen_main_form.dt_feature_codes.Rows[j][1] != DBNull.Value)
                                                        {
                                                            fc = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]);
                                                            if (fc.ToUpper() == feature_code.ToUpper())
                                                            {
                                                                #region check1
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][4] != DBNull.Value)
                                                                {
                                                                    string check1 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][4]);
                                                                    if (check1.Contains("{G}") == true ||
                                                                        check1.Contains("{H}") == true ||
                                                                        check1.Contains("{I}") == true ||
                                                                        check1.Contains("{J}") == true ||
                                                                        check1.Contains("{K}") == true ||
                                                                        check1.Contains("{L}") == true ||
                                                                        check1.Contains("{M}") == true ||
                                                                        check1.Contains("{N}") == true ||
                                                                        check1.Contains("{O}") == true ||
                                                                        check1.Contains("{P}") == true)
                                                                    {
                                                                        if (check1.Contains("{G}") == true)
                                                                        {
                                                                            col_check1 = 6;
                                                                            check1 = check1.Replace("{G}", "");
                                                                            xl_col1 = "G";
                                                                        }

                                                                        if (check1.Contains("{H}") == true)
                                                                        {
                                                                            col_check1 = 7;
                                                                            check1 = check1.Replace("{H}", "");
                                                                            xl_col1 = "H";
                                                                        }

                                                                        if (check1.Contains("{I}") == true)
                                                                        {
                                                                            col_check1 = 8;
                                                                            check1 = check1.Replace("{I}", "");
                                                                            xl_col1 = "I";
                                                                        }

                                                                        if (check1.Contains("{J}") == true)
                                                                        {
                                                                            col_check1 = 9;
                                                                            check1 = check1.Replace("{J}", "");
                                                                            xl_col1 = "J";

                                                                        }

                                                                        if (check1.Contains("{K}") == true)
                                                                        {
                                                                            col_check1 = 10;
                                                                            check1 = check1.Replace("{K}", "");
                                                                            xl_col1 = "K";

                                                                        }

                                                                        if (check1.Contains("{L}") == true)
                                                                        {
                                                                            col_check1 = 11;
                                                                            check1 = check1.Replace("{L}", "");
                                                                            xl_col1 = "L";

                                                                        }

                                                                        if (check1.Contains("{M}") == true)
                                                                        {
                                                                            col_check1 = 12;
                                                                            check1 = check1.Replace("{M}", "");
                                                                            xl_col1 = "M";

                                                                        }

                                                                        if (check1.Contains("{N}") == true)
                                                                        {
                                                                            col_check1 = 13;
                                                                            check1 = check1.Replace("{N}", "");
                                                                            xl_col1 = "N";

                                                                        }

                                                                        if (check1.Contains("{O}") == true)
                                                                        {
                                                                            col_check1 = 14;
                                                                            check1 = check1.Replace("{O}", "");
                                                                            xl_col1 = "O";

                                                                        }

                                                                        if (check1.Contains("{P}") == true)
                                                                        {
                                                                            col_check1 = 15;
                                                                            check1 = check1.Replace("{P}", "");
                                                                            xl_col1 = "P";

                                                                        }

                                                                        if (check1.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check1 = check1.Replace("{NUMBER}", "");
                                                                            number_check1 = true;
                                                                        }

                                                                        check1 = check1.Replace("  ", " ");
                                                                        check1 = check1.Replace("  ", " ");
                                                                        ftype1 = check1;
                                                                        if (ftype1.Length > 0)
                                                                        {
                                                                            if (ftype1.Substring(0, 1) == " ") ftype1 = ftype1.Substring(1, ftype1.Length - 1);
                                                                        }



                                                                    }
                                                                }
                                                                #endregion

                                                                #region check2
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][5] != DBNull.Value)
                                                                {
                                                                    string check2 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][5]);
                                                                    if (check2.Contains("{G}") == true ||
                                                                        check2.Contains("{H}") == true ||
                                                                        check2.Contains("{I}") == true ||
                                                                        check2.Contains("{J}") == true ||
                                                                        check2.Contains("{K}") == true ||
                                                                        check2.Contains("{L}") == true ||
                                                                        check2.Contains("{M}") == true ||
                                                                        check2.Contains("{N}") == true ||
                                                                        check2.Contains("{O}") == true ||
                                                                        check2.Contains("{P}") == true)
                                                                    {
                                                                        if (check2.Contains("{G}") == true)
                                                                        {
                                                                            col_check2 = 6;
                                                                            check2 = check2.Replace("{G}", "");
                                                                            xl_col2 = "G";

                                                                        }

                                                                        if (check2.Contains("{H}") == true)
                                                                        {
                                                                            col_check2 = 7;
                                                                            check2 = check2.Replace("{H}", "");
                                                                            xl_col2 = "H";

                                                                        }

                                                                        if (check2.Contains("{I}") == true)
                                                                        {
                                                                            col_check2 = 8;
                                                                            check2 = check2.Replace("{I}", "");
                                                                            xl_col2 = "I";

                                                                        }

                                                                        if (check2.Contains("{J}") == true)
                                                                        {
                                                                            col_check2 = 9;
                                                                            check2 = check2.Replace("{J}", "");
                                                                            xl_col2 = "J";

                                                                        }

                                                                        if (check2.Contains("{K}") == true)
                                                                        {
                                                                            col_check2 = 10;
                                                                            check2 = check2.Replace("{K}", "");
                                                                            xl_col2 = "K";

                                                                        }

                                                                        if (check2.Contains("{L}") == true)
                                                                        {
                                                                            col_check2 = 11;
                                                                            check2 = check2.Replace("{L}", "");
                                                                            xl_col2 = "L";

                                                                        }

                                                                        if (check2.Contains("{M}") == true)
                                                                        {
                                                                            col_check2 = 12;
                                                                            check2 = check2.Replace("{M}", "");
                                                                            xl_col2 = "M";

                                                                        }

                                                                        if (check2.Contains("{N}") == true)
                                                                        {
                                                                            col_check2 = 13;
                                                                            check2 = check2.Replace("{N}", "");
                                                                            xl_col2 = "N";

                                                                        }

                                                                        if (check2.Contains("{O}") == true)
                                                                        {
                                                                            col_check2 = 14;
                                                                            check2 = check2.Replace("{O}", "");
                                                                            xl_col2 = "O";

                                                                        }

                                                                        if (check2.Contains("{P}") == true)
                                                                        {
                                                                            col_check2 = 15;
                                                                            check2 = check2.Replace("{P}", "");
                                                                            xl_col2 = "P";

                                                                        }

                                                                        if (check2.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check2 = check2.Replace("{NUMBER}", "");
                                                                            number_check2 = true;
                                                                        }

                                                                        check2 = check2.Replace("  ", " ");
                                                                        check2 = check2.Replace("  ", " ");
                                                                        ftype2 = check2;
                                                                        if (ftype2.Length > 0)
                                                                        {
                                                                            if (ftype2.Substring(0, 1) == " ") ftype2 = ftype2.Substring(1, ftype2.Length - 1);
                                                                        }



                                                                    }
                                                                }
                                                                #endregion

                                                                #region check3
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][6] != DBNull.Value)
                                                                {
                                                                    string check3 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][6]);
                                                                    if (check3.Contains("{G}") == true ||
                                                                        check3.Contains("{H}") == true ||
                                                                        check3.Contains("{I}") == true ||
                                                                        check3.Contains("{J}") == true ||
                                                                        check3.Contains("{K}") == true ||
                                                                        check3.Contains("{L}") == true ||
                                                                        check3.Contains("{M}") == true ||
                                                                        check3.Contains("{N}") == true ||
                                                                        check3.Contains("{O}") == true ||
                                                                        check3.Contains("{P}") == true)
                                                                    {
                                                                        if (check3.Contains("{G}") == true)
                                                                        {
                                                                            col_check3 = 6;
                                                                            check3 = check3.Replace("{G}", "");
                                                                            xl_col3 = "G";

                                                                        }

                                                                        if (check3.Contains("{H}") == true)
                                                                        {
                                                                            col_check3 = 7;
                                                                            check3 = check3.Replace("{H}", "");
                                                                            xl_col3 = "H";

                                                                        }

                                                                        if (check3.Contains("{I}") == true)
                                                                        {
                                                                            col_check3 = 8;
                                                                            check3 = check3.Replace("{I}", "");
                                                                            xl_col3 = "I";

                                                                        }

                                                                        if (check3.Contains("{J}") == true)
                                                                        {
                                                                            col_check3 = 9;
                                                                            check3 = check3.Replace("{J}", "");
                                                                            xl_col3 = "J";

                                                                        }

                                                                        if (check3.Contains("{K}") == true)
                                                                        {
                                                                            col_check3 = 10;
                                                                            check3 = check3.Replace("{K}", "");
                                                                            xl_col3 = "K";

                                                                        }

                                                                        if (check3.Contains("{L}") == true)
                                                                        {
                                                                            col_check3 = 11;
                                                                            check3 = check3.Replace("{L}", "");
                                                                            xl_col3 = "L";

                                                                        }

                                                                        if (check3.Contains("{M}") == true)
                                                                        {
                                                                            col_check3 = 12;
                                                                            check3 = check3.Replace("{M}", "");
                                                                            xl_col3 = "M";

                                                                        }

                                                                        if (check3.Contains("{N}") == true)
                                                                        {
                                                                            col_check3 = 13;
                                                                            check3 = check3.Replace("{N}", "");
                                                                            xl_col3 = "N";

                                                                        }

                                                                        if (check3.Contains("{O}") == true)
                                                                        {
                                                                            col_check3 = 14;
                                                                            check3 = check3.Replace("{O}", "");
                                                                            xl_col3 = "O";

                                                                        }

                                                                        if (check3.Contains("{P}") == true)
                                                                        {
                                                                            col_check3 = 15;
                                                                            check3 = check3.Replace("{P}", "");
                                                                            xl_col3 = "P";

                                                                        }

                                                                        if (check3.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check3 = check3.Replace("{NUMBER}", "");
                                                                            number_check3 = true;
                                                                        }

                                                                        check3 = check3.Replace("  ", " ");
                                                                        check3 = check3.Replace("  ", " ");
                                                                        ftype3 = check3;
                                                                        if (ftype3.Length > 0)
                                                                        {
                                                                            if (ftype3.Substring(0, 1) == " ") ftype3 = ftype3.Substring(1, ftype3.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check4
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][7] != DBNull.Value)
                                                                {
                                                                    string check4 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][7]);
                                                                    if (check4.Contains("{G}") == true ||
                                                                        check4.Contains("{H}") == true ||
                                                                        check4.Contains("{I}") == true ||
                                                                        check4.Contains("{J}") == true ||
                                                                        check4.Contains("{K}") == true ||
                                                                        check4.Contains("{L}") == true ||
                                                                        check4.Contains("{M}") == true ||
                                                                        check4.Contains("{N}") == true ||
                                                                        check4.Contains("{O}") == true ||
                                                                        check4.Contains("{P}") == true)
                                                                    {
                                                                        if (check4.Contains("{G}") == true)
                                                                        {
                                                                            col_check4 = 6;
                                                                            check4 = check4.Replace("{G}", "");
                                                                            xl_col4 = "G";

                                                                        }

                                                                        if (check4.Contains("{H}") == true)
                                                                        {
                                                                            col_check4 = 7;
                                                                            check4 = check4.Replace("{H}", "");
                                                                            xl_col4 = "H";

                                                                        }

                                                                        if (check4.Contains("{I}") == true)
                                                                        {
                                                                            col_check4 = 8;
                                                                            check4 = check4.Replace("{I}", "");
                                                                            xl_col4 = "I";

                                                                        }

                                                                        if (check4.Contains("{J}") == true)
                                                                        {
                                                                            col_check4 = 9;
                                                                            check4 = check4.Replace("{J}", "");
                                                                            xl_col4 = "J";

                                                                        }

                                                                        if (check4.Contains("{K}") == true)
                                                                        {
                                                                            col_check4 = 10;
                                                                            check4 = check4.Replace("{K}", "");
                                                                            xl_col4 = "K";

                                                                        }

                                                                        if (check4.Contains("{L}") == true)
                                                                        {
                                                                            col_check4 = 11;
                                                                            check4 = check4.Replace("{L}", "");
                                                                            xl_col4 = "L";

                                                                        }

                                                                        if (check4.Contains("{M}") == true)
                                                                        {
                                                                            col_check4 = 12;
                                                                            check4 = check4.Replace("{M}", "");
                                                                            xl_col4 = "M";

                                                                        }

                                                                        if (check4.Contains("{N}") == true)
                                                                        {
                                                                            col_check4 = 13;
                                                                            check4 = check4.Replace("{N}", "");
                                                                            xl_col4 = "N";

                                                                        }

                                                                        if (check4.Contains("{O}") == true)
                                                                        {
                                                                            col_check4 = 14;
                                                                            check4 = check4.Replace("{O}", "");
                                                                            xl_col4 = "O";
                                                                        }

                                                                        if (check4.Contains("{P}") == true)
                                                                        {
                                                                            col_check4 = 15;
                                                                            check4 = check4.Replace("{P}", "");
                                                                            xl_col4 = "P";
                                                                        }

                                                                        if (check4.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check4 = check4.Replace("{NUMBER}", "");
                                                                            number_check4 = true;
                                                                        }

                                                                        check4 = check4.Replace("  ", " ");
                                                                        check4 = check4.Replace("  ", " ");
                                                                        ftype4 = check4;
                                                                        if (ftype4.Length > 0)
                                                                        {
                                                                            if (ftype4.Substring(0, 1) == " ") ftype4 = ftype4.Substring(1, ftype4.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check5
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][8] != DBNull.Value)
                                                                {
                                                                    string check5 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][8]);
                                                                    if (check5.Contains("{G}") == true ||
                                                                        check5.Contains("{H}") == true ||
                                                                        check5.Contains("{I}") == true ||
                                                                        check5.Contains("{J}") == true ||
                                                                        check5.Contains("{K}") == true ||
                                                                        check5.Contains("{L}") == true ||
                                                                        check5.Contains("{M}") == true ||
                                                                        check5.Contains("{N}") == true ||
                                                                        check5.Contains("{O}") == true ||
                                                                        check5.Contains("{P}") == true)
                                                                    {
                                                                        if (check5.Contains("{G}") == true)
                                                                        {
                                                                            col_check5 = 6;
                                                                            check5 = check5.Replace("{G}", "");
                                                                            xl_col5 = "G";

                                                                        }

                                                                        if (check5.Contains("{H}") == true)
                                                                        {
                                                                            col_check5 = 7;
                                                                            check5 = check5.Replace("{H}", "");
                                                                            xl_col5 = "H";

                                                                        }

                                                                        if (check5.Contains("{I}") == true)
                                                                        {
                                                                            col_check5 = 8;
                                                                            check5 = check5.Replace("{I}", "");
                                                                            xl_col5 = "I";

                                                                        }

                                                                        if (check5.Contains("{J}") == true)
                                                                        {
                                                                            col_check5 = 9;
                                                                            check5 = check5.Replace("{J}", "");
                                                                            xl_col5 = "J";

                                                                        }

                                                                        if (check5.Contains("{K}") == true)
                                                                        {
                                                                            col_check5 = 10;
                                                                            check5 = check5.Replace("{K}", "");
                                                                            xl_col5 = "K";

                                                                        }

                                                                        if (check5.Contains("{L}") == true)
                                                                        {
                                                                            col_check5 = 11;
                                                                            check5 = check5.Replace("{L}", "");
                                                                            xl_col5 = "L";

                                                                        }

                                                                        if (check5.Contains("{M}") == true)
                                                                        {
                                                                            col_check5 = 12;
                                                                            check5 = check5.Replace("{M}", "");
                                                                            xl_col5 = "M";

                                                                        }

                                                                        if (check5.Contains("{N}") == true)
                                                                        {
                                                                            col_check5 = 13;
                                                                            check5 = check5.Replace("{N}", "");
                                                                            xl_col5 = "N";

                                                                        }

                                                                        if (check5.Contains("{O}") == true)
                                                                        {
                                                                            col_check5 = 14;
                                                                            check5 = check5.Replace("{O}", "");
                                                                            xl_col5 = "O";
                                                                        }

                                                                        if (check5.Contains("{P}") == true)
                                                                        {
                                                                            col_check5 = 16;
                                                                            check5 = check5.Replace("{P}", "");
                                                                            xl_col5 = "P";
                                                                        }

                                                                        if (check5.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check5 = check5.Replace("{NUMBER}", "");
                                                                            number_check5 = true;
                                                                        }

                                                                        check5 = check5.Replace("  ", " ");
                                                                        check5 = check5.Replace("  ", " ");
                                                                        ftype5 = check5;
                                                                        if (ftype5.Length > 0)
                                                                        {
                                                                            if (ftype5.Substring(0, 1) == " ") ftype5 = ftype5.Substring(1, ftype5.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check6
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][9] != DBNull.Value)
                                                                {
                                                                    string check6 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][9]);
                                                                    if (check6.Contains("{G}") == true ||
                                                                        check6.Contains("{H}") == true ||
                                                                        check6.Contains("{I}") == true ||
                                                                        check6.Contains("{J}") == true ||
                                                                        check6.Contains("{K}") == true ||
                                                                        check6.Contains("{L}") == true ||
                                                                        check6.Contains("{M}") == true ||
                                                                        check6.Contains("{N}") == true ||
                                                                        check6.Contains("{O}") == true ||
                                                                        check6.Contains("{P}") == true)
                                                                    {
                                                                        if (check6.Contains("{G}") == true)
                                                                        {
                                                                            col_check6 = 6;
                                                                            check6 = check6.Replace("{G}", "");
                                                                            xl_col6 = "G";

                                                                        }

                                                                        if (check6.Contains("{H}") == true)
                                                                        {
                                                                            col_check6 = 7;
                                                                            check6 = check6.Replace("{H}", "");
                                                                            xl_col6 = "H";

                                                                        }

                                                                        if (check6.Contains("{I}") == true)
                                                                        {
                                                                            col_check6 = 8;
                                                                            check6 = check6.Replace("{I}", "");
                                                                            xl_col6 = "I";

                                                                        }

                                                                        if (check6.Contains("{J}") == true)
                                                                        {
                                                                            col_check6 = 9;
                                                                            check6 = check6.Replace("{J}", "");
                                                                            xl_col6 = "J";

                                                                        }

                                                                        if (check6.Contains("{K}") == true)
                                                                        {
                                                                            col_check6 = 10;
                                                                            check6 = check6.Replace("{K}", "");
                                                                            xl_col6 = "K";

                                                                        }

                                                                        if (check6.Contains("{L}") == true)
                                                                        {
                                                                            col_check6 = 11;
                                                                            check6 = check6.Replace("{L}", "");
                                                                            xl_col6 = "L";

                                                                        }

                                                                        if (check6.Contains("{M}") == true)
                                                                        {
                                                                            col_check6 = 12;
                                                                            check6 = check6.Replace("{M}", "");
                                                                            xl_col6 = "M";

                                                                        }

                                                                        if (check6.Contains("{N}") == true)
                                                                        {
                                                                            col_check6 = 13;
                                                                            check6 = check6.Replace("{N}", "");
                                                                            xl_col6 = "N";

                                                                        }

                                                                        if (check6.Contains("{O}") == true)
                                                                        {
                                                                            col_check6 = 14;
                                                                            check6 = check6.Replace("{O}", "");
                                                                            xl_col6 = "O";
                                                                        }

                                                                        if (check6.Contains("{P}") == true)
                                                                        {
                                                                            col_check6 = 15;
                                                                            check6 = check6.Replace("{P}", "");
                                                                            xl_col6 = "P";
                                                                        }

                                                                        if (check6.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check6 = check6.Replace("{NUMBER}", "");
                                                                            number_check6 = true;
                                                                        }

                                                                        check6 = check6.Replace("  ", " ");
                                                                        check6 = check6.Replace("  ", " ");
                                                                        ftype6 = check6;
                                                                        if (ftype2.Length > 0)
                                                                        {
                                                                            if (ftype6.Substring(0, 1) == " ") ftype6 = ftype6.Substring(1, ftype6.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check7
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][10] != DBNull.Value)
                                                                {
                                                                    string check7 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][10]);
                                                                    if (check7.Contains("{G}") == true ||
                                                                        check7.Contains("{H}") == true ||
                                                                        check7.Contains("{I}") == true ||
                                                                        check7.Contains("{J}") == true ||
                                                                        check7.Contains("{K}") == true ||
                                                                        check7.Contains("{L}") == true ||
                                                                        check7.Contains("{M}") == true ||
                                                                        check7.Contains("{N}") == true ||
                                                                        check7.Contains("{O}") == true ||
                                                                        check7.Contains("{P}") == true)
                                                                    {
                                                                        if (check7.Contains("{G}") == true)
                                                                        {
                                                                            col_check7 = 6;
                                                                            check7 = check7.Replace("{G}", "");
                                                                            xl_col7 = "G";

                                                                        }

                                                                        if (check7.Contains("{H}") == true)
                                                                        {
                                                                            col_check7 = 7;
                                                                            check7 = check7.Replace("{H}", "");
                                                                            xl_col7 = "H";

                                                                        }

                                                                        if (check7.Contains("{I}") == true)
                                                                        {
                                                                            col_check7 = 8;
                                                                            check7 = check7.Replace("{I}", "");
                                                                            xl_col7 = "I";

                                                                        }

                                                                        if (check7.Contains("{J}") == true)
                                                                        {
                                                                            col_check7 = 9;
                                                                            check7 = check7.Replace("{J}", "");
                                                                            xl_col7 = "J";

                                                                        }

                                                                        if (check7.Contains("{K}") == true)
                                                                        {
                                                                            col_check7 = 10;
                                                                            check7 = check7.Replace("{K}", "");
                                                                            xl_col7 = "K";

                                                                        }

                                                                        if (check7.Contains("{L}") == true)
                                                                        {
                                                                            col_check7 = 11;
                                                                            check7 = check7.Replace("{L}", "");
                                                                            xl_col7 = "L";

                                                                        }

                                                                        if (check7.Contains("{M}") == true)
                                                                        {
                                                                            col_check7 = 12;
                                                                            check7 = check7.Replace("{M}", "");
                                                                            xl_col7 = "M";

                                                                        }

                                                                        if (check7.Contains("{N}") == true)
                                                                        {
                                                                            col_check7 = 13;
                                                                            check7 = check7.Replace("{N}", "");
                                                                            xl_col7 = "N";

                                                                        }

                                                                        if (check7.Contains("{O}") == true)
                                                                        {
                                                                            col_check7 = 14;
                                                                            check7 = check7.Replace("{O}", "");
                                                                            xl_col7 = "O";
                                                                        }

                                                                        if (check7.Contains("{P}") == true)
                                                                        {
                                                                            col_check7 = 15;
                                                                            check7 = check7.Replace("{P}", "");
                                                                            xl_col7 = "P";
                                                                        }

                                                                        if (check7.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check7 = check7.Replace("{NUMBER}", "");
                                                                            number_check7 = true;
                                                                        }

                                                                        check7 = check7.Replace("  ", " ");
                                                                        check7 = check7.Replace("  ", " ");
                                                                        ftype7 = check7;
                                                                        if (ftype7.Length > 0)
                                                                        {
                                                                            if (ftype7.Substring(0, 1) == " ") ftype7 = ftype7.Substring(1, ftype7.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check8
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][11] != DBNull.Value)
                                                                {
                                                                    string check8 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][11]);
                                                                    if (check8.Contains("{G}") == true ||
                                                                        check8.Contains("{H}") == true ||
                                                                        check8.Contains("{I}") == true ||
                                                                        check8.Contains("{J}") == true ||
                                                                        check8.Contains("{K}") == true ||
                                                                        check8.Contains("{L}") == true ||
                                                                        check8.Contains("{M}") == true ||
                                                                        check8.Contains("{N}") == true ||
                                                                        check8.Contains("{O}") == true ||
                                                                        check8.Contains("{P}") == true)
                                                                    {
                                                                        if (check8.Contains("{G}") == true)
                                                                        {
                                                                            col_check8 = 6;
                                                                            check8 = check8.Replace("{G}", "");
                                                                            xl_col8 = "G";

                                                                        }

                                                                        if (check8.Contains("{H}") == true)
                                                                        {
                                                                            col_check8 = 7;
                                                                            check8 = check8.Replace("{H}", "");
                                                                            xl_col8 = "H";

                                                                        }

                                                                        if (check8.Contains("{I}") == true)
                                                                        {
                                                                            col_check8 = 8;
                                                                            check8 = check8.Replace("{I}", "");
                                                                            xl_col8 = "I";

                                                                        }

                                                                        if (check8.Contains("{J}") == true)
                                                                        {
                                                                            col_check8 = 9;
                                                                            check8 = check8.Replace("{J}", "");
                                                                            xl_col8 = "J";

                                                                        }

                                                                        if (check8.Contains("{K}") == true)
                                                                        {
                                                                            col_check8 = 10;
                                                                            check8 = check8.Replace("{K}", "");
                                                                            xl_col8 = "K";

                                                                        }

                                                                        if (check8.Contains("{L}") == true)
                                                                        {
                                                                            col_check8 = 11;
                                                                            check8 = check8.Replace("{L}", "");
                                                                            xl_col8 = "L";

                                                                        }

                                                                        if (check8.Contains("{M}") == true)
                                                                        {
                                                                            col_check8 = 12;
                                                                            check8 = check8.Replace("{M}", "");
                                                                            xl_col8 = "M";

                                                                        }

                                                                        if (check8.Contains("{N}") == true)
                                                                        {
                                                                            col_check8 = 13;
                                                                            check8 = check8.Replace("{N}", "");
                                                                            xl_col8 = "N";

                                                                        }

                                                                        if (check8.Contains("{O}") == true)
                                                                        {
                                                                            col_check8 = 14;
                                                                            check8 = check8.Replace("{O}", "");
                                                                            xl_col8 = "O";
                                                                        }

                                                                        if (check8.Contains("{P}") == true)
                                                                        {
                                                                            col_check8 = 15;
                                                                            check8 = check8.Replace("{P}", "");
                                                                            xl_col8 = "P";
                                                                        }

                                                                        if (check8.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check8 = check8.Replace("{NUMBER}", "");
                                                                            number_check8 = true;
                                                                        }

                                                                        check8 = check8.Replace("  ", " ");
                                                                        check8 = check8.Replace("  ", " ");
                                                                        ftype8 = check8;

                                                                        if (ftype8.Length > 0)
                                                                        {
                                                                            if (ftype8.Substring(0, 1) == " ") ftype8 = ftype8.Substring(1, ftype8.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check9
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][12] != DBNull.Value)
                                                                {
                                                                    string check9 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][12]);
                                                                    if (check9.Contains("{G}") == true ||
                                                                        check9.Contains("{H}") == true ||
                                                                        check9.Contains("{I}") == true ||
                                                                        check9.Contains("{J}") == true ||
                                                                        check9.Contains("{K}") == true ||
                                                                        check9.Contains("{L}") == true ||
                                                                        check9.Contains("{M}") == true ||
                                                                        check9.Contains("{N}") == true ||
                                                                        check9.Contains("{O}") == true ||
                                                                        check9.Contains("{P}") == true)
                                                                    {
                                                                        if (check9.Contains("{G}") == true)
                                                                        {
                                                                            col_check9 = 6;
                                                                            check9 = check9.Replace("{G}", "");
                                                                            xl_col9 = "G";

                                                                        }

                                                                        if (check9.Contains("{H}") == true)
                                                                        {
                                                                            col_check9 = 7;
                                                                            check9 = check9.Replace("{H}", "");
                                                                            xl_col9 = "H";

                                                                        }

                                                                        if (check9.Contains("{I}") == true)
                                                                        {
                                                                            col_check9 = 8;
                                                                            check9 = check9.Replace("{I}", "");
                                                                            xl_col9 = "I";

                                                                        }

                                                                        if (check9.Contains("{J}") == true)
                                                                        {
                                                                            col_check9 = 9;
                                                                            check9 = check9.Replace("{J}", "");
                                                                            xl_col9 = "J";

                                                                        }

                                                                        if (check9.Contains("{K}") == true)
                                                                        {
                                                                            col_check9 = 10;
                                                                            check9 = check9.Replace("{K}", "");
                                                                            xl_col9 = "K";

                                                                        }

                                                                        if (check9.Contains("{L}") == true)
                                                                        {
                                                                            col_check9 = 11;
                                                                            check9 = check9.Replace("{L}", "");
                                                                            xl_col9 = "L";

                                                                        }

                                                                        if (check9.Contains("{M}") == true)
                                                                        {
                                                                            col_check9 = 12;
                                                                            check9 = check9.Replace("{M}", "");
                                                                            xl_col9 = "M";

                                                                        }

                                                                        if (check9.Contains("{N}") == true)
                                                                        {
                                                                            col_check9 = 13;
                                                                            check9 = check9.Replace("{N}", "");
                                                                            xl_col9 = "N";

                                                                        }

                                                                        if (check9.Contains("{O}") == true)
                                                                        {
                                                                            col_check9 = 14;
                                                                            check9 = check9.Replace("{O}", "");
                                                                            xl_col9 = "O";
                                                                        }

                                                                        if (check9.Contains("{P}") == true)
                                                                        {
                                                                            col_check9 = 15;
                                                                            check9 = check9.Replace("{P}", "");
                                                                            xl_col9 = "P";
                                                                        }

                                                                        if (check9.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check9 = check9.Replace("{NUMBER}", "");
                                                                            number_check9 = true;
                                                                        }

                                                                        check9 = check9.Replace("  ", " ");
                                                                        check9 = check9.Replace("  ", " ");
                                                                        ftype9 = check9;
                                                                        if (ftype9.Length > 0)
                                                                        {
                                                                            if (ftype9.Substring(0, 1) == " ") ftype9 = ftype9.Substring(1, ftype9.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                #region check10
                                                                if (Wgen_main_form.dt_feature_codes.Rows[j][13] != DBNull.Value)
                                                                {
                                                                    string check10 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][13]);
                                                                    if (check10.Contains("{G}") == true ||
                                                                        check10.Contains("{H}") == true ||
                                                                        check10.Contains("{I}") == true ||
                                                                        check10.Contains("{J}") == true ||
                                                                        check10.Contains("{K}") == true ||
                                                                        check10.Contains("{L}") == true ||
                                                                        check10.Contains("{M}") == true ||
                                                                        check10.Contains("{N}") == true ||
                                                                        check10.Contains("{O}") == true ||
                                                                        check10.Contains("{P}") == true)
                                                                    {
                                                                        if (check10.Contains("{G}") == true)
                                                                        {
                                                                            col_check10 = 6;
                                                                            check10 = check10.Replace("{G}", "");
                                                                            xl_col10 = "G";

                                                                        }

                                                                        if (check10.Contains("{H}") == true)
                                                                        {
                                                                            col_check10 = 7;
                                                                            check10 = check10.Replace("{H}", "");
                                                                            xl_col10 = "H";

                                                                        }

                                                                        if (check10.Contains("{I}") == true)
                                                                        {
                                                                            col_check10 = 8;
                                                                            check10 = check10.Replace("{I}", "");
                                                                            xl_col10 = "I";

                                                                        }

                                                                        if (check10.Contains("{J}") == true)
                                                                        {
                                                                            col_check10 = 9;
                                                                            check10 = check10.Replace("{J}", "");
                                                                            xl_col10 = "J";

                                                                        }

                                                                        if (check10.Contains("{K}") == true)
                                                                        {
                                                                            col_check10 = 10;
                                                                            check10 = check10.Replace("{K}", "");
                                                                            xl_col10 = "K";

                                                                        }

                                                                        if (check10.Contains("{L}") == true)
                                                                        {
                                                                            col_check10 = 11;
                                                                            check10 = check10.Replace("{L}", "");
                                                                            xl_col10 = "L";

                                                                        }

                                                                        if (check10.Contains("{M}") == true)
                                                                        {
                                                                            col_check10 = 12;
                                                                            check10 = check10.Replace("{M}", "");
                                                                            xl_col10 = "M";

                                                                        }

                                                                        if (check10.Contains("{N}") == true)
                                                                        {
                                                                            col_check10 = 13;
                                                                            check10 = check10.Replace("{N}", "");
                                                                            xl_col10 = "N";

                                                                        }

                                                                        if (check10.Contains("{O}") == true)
                                                                        {
                                                                            col_check10 = 14;
                                                                            check10 = check10.Replace("{O}", "");
                                                                            xl_col10 = "O";
                                                                        }

                                                                        if (check10.Contains("{P}") == true)
                                                                        {
                                                                            col_check10 = 15;
                                                                            check10 = check10.Replace("{P}", "");
                                                                            xl_col10 = "P";
                                                                        }

                                                                        if (check10.Contains("{NUMBER}") == true)
                                                                        {
                                                                            check10 = check10.Replace("{NUMBER}", "");
                                                                            number_check10 = true;
                                                                        }

                                                                        check10 = check10.Replace("  ", " ");
                                                                        check10 = check10.Replace("  ", " ");
                                                                        ftype10 = check10;
                                                                        if (ftype10.Length > 0)
                                                                        {
                                                                            if (ftype10.Substring(0, 1) == " ") ftype10 = ftype10.Substring(1, ftype10.Length - 1);
                                                                        }

                                                                    }
                                                                }
                                                                #endregion

                                                                j = Wgen_main_form.dt_feature_codes.Rows.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            #region check1
                                            if (col_check1 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check1] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col1 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype1 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }
                                            if (col_check1 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check1] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check1])) == false && number_check1 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;

                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col1 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype1 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check2
                                            if (col_check2 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check2] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col2 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype2 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check2 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check2] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check2])) == false && number_check2 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col2 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype2 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check3
                                            if (col_check3 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check3] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col3 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype3 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check3 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check3] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check3])) == false && number_check3 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col3 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype3 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check4
                                            if (col_check4 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check4] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col4 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype4 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check4 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check4] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check4])) == false && number_check4 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col4 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype4 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check5
                                            if (col_check5 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check5] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col5 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype5 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check5 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check5] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check5])) == false && number_check5 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col5 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype5 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check6
                                            if (col_check6 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check6] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col6 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype6 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check6 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check6] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check6])) == false && number_check6 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col6 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype6 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check7
                                            if (col_check7 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check7] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col7 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype7 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check7 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check7] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check7])) == false && number_check7 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col7 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype7 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check8
                                            if (col_check8 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check8] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col8 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype8 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check9 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check8] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check8])) == false && number_check8 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col8 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype8 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check9
                                            if (col_check9 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check9] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col9 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype9 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check9 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check9] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check9])) == false && number_check9 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col9 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype9 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                            #region check10
                                            if (col_check10 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check10] == DBNull.Value)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col10 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = "No " + fc + " " + ftype10 + " " + "Specified";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                                ++nr_null_values;
                                            }

                                            if (col_check10 > 0 && Wgen_main_form.dt_all_points.Rows[i][col_check10] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_check10])) == false && number_check10 == true)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Point"] = pt1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Excel"] = xl_col10 + Convert.ToString(Convert.ToInt32(Wgen_main_form.dt_all_points.Rows[i]["index1"]) + start_row);
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["w1"] = W1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error"] = fc + " " + ftype10 + " not numeric";
                                                string x = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col3] != DBNull.Value)
                                                {
                                                    x = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col3]);
                                                }
                                                string y = "";
                                                if (Wgen_main_form.dt_all_points.Rows[i][col2] != DBNull.Value)
                                                {
                                                    y = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col2]);
                                                }

                                                dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;
                                            }
                                            #endregion

                                        }
                                    }
                                    #endregion


                                }




                                dt_errors = Functions.Sort_data_table(dt_errors, "Error");
                                transfer_errors_to_panel(dt_errors);
                                dt_export = Functions.creaza_error_export_table(dt_errors, sheet_name);
                                textBox_AP_no_rows.Text = Convert.ToString(Wgen_main_form.dt_all_points.Rows.Count);
                                textBox_AP_no_duplicates.Text = Convert.ToString(nr_duplicates);
                                textBox_AP_no_null.Text = Convert.ToString(nr_null_values);
                                textBox_AP_no_mmid.Text = Convert.ToString(nr_mmid_not_found);
                                textBox_AP_no_xray.Text = Convert.ToString(nr_xray_duplicates);

                                button_all_pts_l.Visible = true;
                                button_all_pts_nl.Visible = false;
                            }
                            else
                            {
                                button_all_pts_l.Visible = false;
                                button_all_pts_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_all_pts_l.Visible = false;
                        button_all_pts_nl.Visible = true;
                    }
                }
                else
                {
                    button_all_pts_l.Visible = false;
                    button_all_pts_nl.Visible = true;
                }
            }
            else
            {
                button_all_pts_l.Visible = false;
                button_all_pts_nl.Visible = true;
            }

            // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general( Wgen_main_form. dt_all_points);
        }

      
        private void button_zoom_click(object sender, EventArgs e)
        {

            if (dt_errors == null || dt_errors.Rows.Count == 0) return;

            int index1 = 0;
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet W1 = (Microsoft.Office.Interop.Excel.Worksheet)dt_errors.Rows[index1]["w1"];
                if (W1 != null)
                {
                    if (dt_errors.Rows[index1]["Excel"] != DBNull.Value)
                    {
                        string adresa = Convert.ToString(dt_errors.Rows[index1]["Excel"]);
                        W1.Activate();
                        W1.Range[adresa].Select();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void make_first_line_invisible()
        {
            dt_display = new System.Data.DataTable();
            dt_display.Columns.Add("Point", typeof(string));
            dt_display.Columns.Add("Value", typeof(string));
            dt_display.Columns.Add("Excel", typeof(string));
            dt_display.Columns.Add("Error", typeof(string));
            dataGridView_error_all_pts.DataSource = dt_display;
            dataGridView_error_all_pts.Columns[0].Width = 75;
            dataGridView_error_all_pts.Columns[1].Width = 75;
            dataGridView_error_all_pts.Columns[2].Width = 50;
            dataGridView_error_all_pts.Columns[3].Width = 300;
            dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_error_all_pts.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_error_all_pts.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_error_all_pts.EnableHeadersVisualStyles = false;
        }



        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            if (comboBox_ws1.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
                {
                    if (comboBox_ws1.Items[i].ToString().ToUpper().Contains("ASBUILT_POINTS") == true)
                    {
                        comboBox_ws1.SelectedIndex = i;
                        i = comboBox_ws1.Items.Count;
                    }
                }
            }

        }

        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_named(dt_export, "AllPointsErrors");
        }

        private void button_load_features_Click(object sender, EventArgs e)
        {
            foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
            {
                if (Forma1 is Alignment_mdi.Wgen_feature)
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
                Alignment_mdi.Wgen_feature forma2 = new Alignment_mdi.Wgen_feature();
                Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                     (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);
            }
            catch (System.Exception EX)
            {
                MessageBox.Show(EX.Message);
            }
        }

        private void DataGridView_error_all_pts_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_error_all_pts.CurrentCell = dataGridView_error_all_pts.Rows[e.RowIndex].Cells[0];
                ContextMenuStrip_go_to_error.Show(Cursor.Position);
                ContextMenuStrip_go_to_error.Visible = true;
            }
            else
            {
                ContextMenuStrip_go_to_error.Visible = false;
            }
        }

        private void go_to_excel_point(object sender, EventArgs e)
        {
            if (dt_errors == null || dt_errors.Rows.Count == 0) return;

            int index1 = dataGridView_error_all_pts.CurrentCell.RowIndex;
            try
            {
                if (dt_errors.Rows.Count - 1 >= index1)
                {
                    Microsoft.Office.Interop.Excel.Worksheet W1 = (Microsoft.Office.Interop.Excel.Worksheet)dt_errors.Rows[index1]["w1"];
                    if (W1 != null)
                    {
                        if (dt_errors.Rows[index1]["Excel"] != DBNull.Value)
                        {
                            string adresa = Convert.ToString(dt_errors.Rows[index1]["Excel"]);
                            W1.Activate();
                            W1.Range[adresa].Select();
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        private void zoom_to_point_in_acad(object sender, EventArgs e)
        {
            if (dt_errors == null || dt_errors.Rows.Count == 0) return;

            int index1 = dataGridView_error_all_pts.CurrentCell.RowIndex;
            try
            {
                if (dt_errors.Rows.Count - 1 >= index1)
                {
                    if (dt_errors.Rows[index1][5] != DBNull.Value && dt_errors.Rows[index1][6] != DBNull.Value
                        && Functions.IsNumeric(Convert.ToString(dt_errors.Rows[index1][5])) == true && Functions.IsNumeric(Convert.ToString(dt_errors.Rows[index1][6])) == true)
                    {
                        double x = Convert.ToDouble(dt_errors.Rows[index1][5]);
                        double y = Convert.ToDouble(dt_errors.Rows[index1][6]);

                        Functions.zoom_to_Point(new Point3d(x, y, 0), 5);

                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }



        }

        public void radioButton_all_points_CheckedChanged(RadioButton radioButton_enlarged)
        {

           System.Drawing. Font regularfont = new Font("Arial", 8.2f, FontStyle.Bold);


            Font englargedFont = new Font("Arial", 10f, FontStyle.Bold);


            Font regularHeader = new Font("Arial", 10f, FontStyle.Bold);

            Font englargedHeader = new Font("Arial", 12f, FontStyle.Bold);

            if (radioButton_enlarged.Checked == true)
            {
                label12.Location = new Point(5, 5);
                label12.Size = new Size(93, 23);

                label12.Font = englargedHeader;
            }
            else
            {
                label12.Location = new Point(3, 3);
                label12.Size = new Size(75, 18);

                label12.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel7.Location = new Point(0, 0);
                panel7.Size = new Size(723, 28);
            }
            else
            {
                panel7.Location = new Point(0, 0);
                panel7.Size = new Size(723, 25);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label25.Location = new Point(150, 5);
                label25.Size = new Size(65, 18);

                label25.Font = englargedHeader;
            }
            else
            {
                label25.Location = new Point(150, 3);
                label25.Size = new Size(65, 18);

                label25.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                label_client.Location = new Point(221, 5);
                label_client.Size = new Size(65, 18);

                label_client.Font = englargedHeader;
            }
            else
            {
                label_client.Location = new Point(221, 3);
                label_client.Size = new Size(65, 18);

                label_client.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_pipe_manifest.Location = new Point(3, 30);
                panel_pipe_manifest.Size = new Size(723, 35);
            }
            else
            {
                panel_pipe_manifest.Location = new Point(3, 27);
                panel_pipe_manifest.Size = new Size(723, 33);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_all_pts_l.Location = new Point(695, 3);
                button_all_pts_l.Size = new Size(24, 24);
            }
            else
            {
                button_all_pts_l.Location = new Point(697, 5);
                button_all_pts_l.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_all_pts_nl.Location = new Point(695, 3);
                button_all_pts_nl.Size = new Size(24, 24);
            }
            else
            {
                button_all_pts_nl.Location = new Point(697, 5);
                button_all_pts_nl.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_refresh_ws1.Location = new Point(3, 3);
                button_refresh_ws1.Size = new Size(24, 24);
            }
            else
            {
                button_refresh_ws1.Location = new Point(5, 5);
                button_refresh_ws1.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_load_all_points.Location = new Point(494, 2);
                button_load_all_points.Size = new Size(198, 30);

                button_load_all_points.Font = englargedFont;
            }
            else
            {
                button_load_all_points.Location = new Point(546, 2);
                button_load_all_points.Size = new Size(145, 28);

                button_load_all_points.Font = regularfont;
            }


            if (radioButton_enlarged.Checked == true)
            {
                comboBox_ws1.Location = new Point(32, 4);
                comboBox_ws1.Size = new Size(455, 25);

                comboBox_ws1.Font = englargedFont;
            }
            else
            {
                comboBox_ws1.Location = new Point(32, 4);
                comboBox_ws1.Size = new Size(508, 24);

                comboBox_ws1.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                label18.Location = new Point(3, 3);
                label18.Size = new Size(81, 30);

                label18.Font = englargedHeader;
            }
            else
            {
                label18.Location = new Point(3, 3);
                label18.Size = new Size(81, 25);

                label18.Font = regularHeader;
            }





            if (radioButton_enlarged.Checked == true)
            {
                panel_stats.Location = new Point(3, 520);
                panel_stats.Size = new Size(722, 152);
            }
            else
            {
                panel_stats.Location = new Point(3, 542);
                panel_stats.Size = new Size(722, 128);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label19.Location = new Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = englargedFont;
            }
            else
            {
                label19.Location = new Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_Items.Location = new Point(3, 21);
                textBox_AP_Items.Size = new Size(300, 25);

                textBox_AP_Items.Font = englargedFont;
            }
            else
            {
                textBox_AP_Items.Location = new Point(3, 21);
                textBox_AP_Items.Size = new Size(287, 20);

                textBox_AP_Items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_no_rows.Location = new Point(315, 21);
                textBox_AP_no_rows.Size = new Size(40, 25);

                textBox_AP_no_rows.Font = englargedFont;
            }
            else
            {
                textBox_AP_no_rows.Location = new Point(311, 21);
                textBox_AP_no_rows.Size = new Size(37, 20);

                textBox_AP_no_rows.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_missing_OD.Location = new Point(3, 47);
                textBox_AP_missing_OD.Size = new Size(300, 25);

                textBox_AP_missing_OD.Font = englargedFont;
            }
            else
            {
                textBox_AP_missing_OD.Location = new Point(3, 42);
                textBox_AP_missing_OD.Size = new Size(287, 20);

                textBox_AP_missing_OD.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_no_duplicates.Location = new Point(315, 47);
                textBox_AP_no_duplicates.Size = new Size(40, 25);

                textBox_AP_no_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_AP_no_duplicates.Location = new Point(311, 42);
                textBox_AP_no_duplicates.Size = new Size(37, 20);

                textBox_AP_no_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_MMID.Location = new Point(3, 73);
                textBox_AP_MMID.Size = new Size(300, 25);

                textBox_AP_MMID.Font = englargedFont;
            }
            else
            {
                textBox_AP_MMID.Location = new Point(3, 63);
                textBox_AP_MMID.Size = new Size(287, 20);

                textBox_AP_MMID.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_no_mmid.Location = new Point(315, 73);
                textBox_AP_no_mmid.Size = new Size(40, 25);

                textBox_AP_no_mmid.Font = englargedFont;
            }
            else
            {
                textBox_AP_no_mmid.Location = new Point(311, 63);
                textBox_AP_no_mmid.Size = new Size(37, 20);

                textBox_AP_no_mmid.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_Xray.Location = new Point(3, 99);
                textBox_AP_Xray.Size = new Size(300, 25);

                textBox_AP_Xray.Font = englargedFont;
            }
            else
            {
                textBox_AP_Xray.Location = new Point(3, 84);
                textBox_AP_Xray.Size = new Size(287, 20);

                textBox_AP_Xray.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_no_xray.Location = new Point(315, 99);
                textBox_AP_no_xray.Size = new Size(40, 25);

                textBox_AP_no_xray.Font = englargedFont;
            }
            else
            {
                textBox_AP_no_xray.Location = new Point(311, 84);
                textBox_AP_no_xray.Size = new Size(37, 20);

                textBox_AP_no_xray.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_null_value_items.Location = new Point(3, 125);
                textBox_AP_null_value_items.Size = new Size(300, 25);

                textBox_AP_null_value_items.Font = englargedFont;
            }
            else
            {
                textBox_AP_null_value_items.Location = new Point(3, 105);
                textBox_AP_null_value_items.Size = new Size(287, 20);

                textBox_AP_null_value_items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_AP_no_null.Location = new Point(315, 125);
                textBox_AP_no_null.Size = new Size(40, 25);

                textBox_AP_no_null.Font = englargedFont;
            }
            else
            {
                textBox_AP_no_null.Location = new Point(311, 105);
                textBox_AP_no_null.Size = new Size(37, 20);

                textBox_AP_no_null.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_export_errors_to_xl.Location = new Point(504, 118);
                button_export_errors_to_xl.Size = new Size(214, 30);

                button_export_errors_to_xl.Font = englargedFont;
            }
            else
            {
                button_export_errors_to_xl.Location = new Point(556, 94);
                button_export_errors_to_xl.Size = new Size(161, 28);

                button_export_errors_to_xl.Font = regularfont;
            }


            if (radioButton_enlarged.Checked == true)
            {
                dataGridView_error_all_pts.Location = new Point(2, 97);
                dataGridView_error_all_pts.Size = new Size(723, 420);

                dataGridView_error_all_pts.DefaultCellStyle.Font = englargedFont;

                dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.Font = englargedHeader;

            }
            else
            {
                dataGridView_error_all_pts.Location = new Point(2, 91);
                dataGridView_error_all_pts.Size = new Size(723, 450);

                dataGridView_error_all_pts.DefaultCellStyle.Font = regularfont;
                dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_dismiss_errors.Location = new Point(386, 0);
                button_dismiss_errors.Size = new Size(124, 30);

                button_dismiss_errors.Font = englargedFont;
            }
            else
            {
                button_dismiss_errors.Location = new Point(437,0);
                button_dismiss_errors.Size = new Size(104, 25);

                button_dismiss_errors.Font = regularfont;
            }


            if (radioButton_enlarged.Checked == true)
            {
                panel6.Location = new Point(2, 69);
                panel6.Size = new Size(723, 35);
            }
            else
            {
                panel6.Location = new Point(2, 64);
                panel6.Size = new Size(723, 30);
            }

        }

        private void transfer_errors_to_panel(System.Data.DataTable dt1)
        {

            //dt_errors = new System.Data.DataTable();
            //dt_errors.Columns.Add("Point", typeof(string));
            //dt_errors.Columns.Add("Value", typeof(string));
            //dt_errors.Columns.Add("Excel", typeof(string));
            //dt_errors.Columns.Add("w1", typeof(Microsoft.Office.Interop.Excel.Worksheet));
            //dt_errors.Columns.Add("Error", typeof(string));
            //dt_errors.Columns.Add("x", typeof(string));
            //dt_errors.Columns.Add("y", typeof(string));

            if (dt1.Rows.Count > 0)
            {

                dt_display = dt1.Copy();
                dt_display.Columns.RemoveAt(6);
                dt_display.Columns.RemoveAt(5);
                dt_display.Columns.RemoveAt(3);

                if (dt_dismissed_errors != null && dt_dismissed_errors.Rows.Count > 0)
                {
                    dt_display.TableName = "err1";
                    dt_dismissed_errors.TableName = "t5";
                    DataSet dataset1 = new DataSet();
                    dataset1.Tables.Add(dt_dismissed_errors);
                    dataset1.Tables.Add(dt_display);

                    DataRelation rel0 = new DataRelation("xxx", dt_display.Columns[0], dt_dismissed_errors.Columns[0], false);
                    DataRelation rel1 = new DataRelation("xxx1", dt_display.Columns[1], dt_dismissed_errors.Columns[1], false);
                    DataRelation rel2 = new DataRelation("xxx2", dt_display.Columns[3], dt_dismissed_errors.Columns[2], false);


                    dataset1.Relations.Add(rel0);
                    dataset1.Relations.Add(rel1);
                    dataset1.Relations.Add(rel2);

                    for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
                    {
                        if (dt_display.Rows[i][0] != DBNull.Value && dt_display.Rows[i][1] != DBNull.Value && dt_display.Rows[i][3] != DBNull.Value)
                        {
                            if (dt_display.Rows[i].GetChildRows(rel0).Length > 0 && dt_display.Rows[i].GetChildRows(rel1).Length > 0 && dt_display.Rows[i].GetChildRows(rel2).Length > 0)
                            {
                                // the 3 rel are not the right way... they are taking out extra lines
                                dt_display.Rows[i].Delete();
                            }
                        }
                    }

                    dataset1.Relations.Remove(rel0);
                    dataset1.Relations.Remove(rel1);
                    dataset1.Relations.Remove(rel2);
                    dataset1.Tables.Remove(dt_dismissed_errors);
                    dataset1.Tables.Remove(dt_display);


                }

                dataGridView_error_all_pts.DataSource = dt_display;
                dataGridView_error_all_pts.Columns[0].Width = 75;
                dataGridView_error_all_pts.Columns[1].Width = 75;
                dataGridView_error_all_pts.Columns[2].Width = 50;
                dataGridView_error_all_pts.Columns[3].Width = 300;
                dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_all_pts.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_all_pts.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_all_pts.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_all_pts.EnableHeadersVisualStyles = false;
            }
        }

        private void button_dismiss_errors_Click(object sender, EventArgs e)
        {
            try
            {
                if (dt_dismissed_errors == null)
                {
                    dt_dismissed_errors = new System.Data.DataTable();
                    dt_dismissed_errors.Columns.Add("Point", typeof(string));
                    dt_dismissed_errors.Columns.Add("Value", typeof(string));
                    dt_dismissed_errors.Columns.Add("Error", typeof(string));
                }

                List<int> lista1 = new List<int>();

                foreach (DataGridViewCell cell1 in dataGridView_error_all_pts.SelectedCells)
                {
                    int row_index = cell1.RowIndex;
                    if (lista1.Contains(row_index) == false)
                    {
                        lista1.Add(row_index);
                    }
                }


                if (lista1.Count > 0)
                {
                    for (int i = 0; i < lista1.Count; ++i)
                    {
                        dt_dismissed_errors.Rows.Add();
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][0] = dataGridView_error_all_pts.Rows[lista1[i]].Cells[0].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][1] = dataGridView_error_all_pts.Rows[lista1[i]].Cells[1].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][2] = dataGridView_error_all_pts.Rows[lista1[i]].Cells[3].Value;
                    }

                    if (W2 == null)
                    {
                        Functions.Create_a_new_worksheet_from_excel_by_name(filename, dismiss_errors_tab);

                    }
                    Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(dt_dismissed_errors, filename, dismiss_errors_tab, false);

                    transfer_errors_to_panel(dt_errors);

                }


            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
    }
}
