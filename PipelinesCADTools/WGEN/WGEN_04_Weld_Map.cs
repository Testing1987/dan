using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Font = System.Drawing.Font;


namespace Alignment_mdi
{
    public partial class Wgen_weldmap : Form
    {
        private ContextMenuStrip ContextMenuStrip_go_to_error;

        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_st_eq;

        System.Data.DataTable dt_welds_with_pmc = null;


        double ng_tolerance = 100;
        double length_check_tolerance = 1;
        double pmc_tolerance = 1;
        System.Windows.Forms.TextBox textbox1 = null;

        Microsoft.Office.Interop.Excel.Worksheet W2 = null;

        int start_row = 2;
        string col_pnt = "PNT";
        string col_y = "NORTHING";
        string col_x = "EASTING";
        string col_z = "ELEVATION";
        string col_feat_code = "FEATURE_CODE";
        string col_descr = "DESCRIPTION";
        string col_sta_lin = "STATION_LINEAR";
        string col_sta_ifc = "STATION_IFC";
        string col_mm_bk = "MM_BK";
        string col_wall_bk = "WALL_BK";
        string col_pipe_bk = "PIPE_BK";
        string col_heat_bk = "HEAT_BK";
        string col_coat_bk = "COATING_BK";
        string col_grade_bk = "GRADE_BK";
        string col_mm_ahd = "MM_AHD";
        string col_wall_ahd = "WALL_AHD";
        string col_pipe_ahd = "PIPE_AHD";
        string col_heat_ahd = "HEAT_AHD";
        string col_coat_ahd = "COATING_AHD";
        string col_grade_ahd = "GRADE_AHD";
        string col_ng = "NG";
        string col_ng_y = "NG_NORTHING";
        string col_ng_x = "NG_EASTING";
        string col_ng_z = "NG_ELEVATION";
        string col_cover = "COVER";
        string col_location = "LOCATION";
        string col_file_name = "FILENAME";
        string col_length_bk = "LENGTH BACK";
        string col_length_ahd = "LENGTH AHEAD";
        string col_manufacture_bk = "MANUFACTURER BACK";
        string col_manufacture_ahd = "MANUFACTURER AHEAD";



        string colpt1 = "PNT";
        string colpt2 = "NORTHING";
        string colpt3 = "EASTING";
        string colpt4 = "ELEVATION";
        string colpt5 = "FEATURE CODE";
        string colpt6 = "FILENAME";
        string colpt7 = "LOCATION";
        string colpt8 = "NOTES";
        string colpt9 = "DESCRIPTION";
        string colpt10 = "MISC1";
        string colpt11 = "MISC2";
        string colpt12 = "MISC3";
        string colpt13 = "MISC4";
        string colpt14 = "MISC5";
        string colpt15 = "MISC6";
        string colpt16 = "MISC7";
        string colpt17 = "MISC8";
        string colpt18 = "MISC9";
        string colpt19 = "MISC10";
        string colpt20 = "MISC11";
        string colpt21 = "MISC12";
        string colpt22 = "MISC13";
        string colpt23 = "MISC14";
        string colpt24 = "MISC15";
        string colpt25 = "MISC16";
        string colpt26 = "MISC17";
        string colpt_sta = "STATION_LINEAR";
        string colpt28 = "STATION_IFC";


        string colgt1 = "MMID";
        string colgt2 = "Pipe";
        string colgt3 = "Heat";
        string colgt4 = "OriginalLength";
        string colgt5 = "NewLength";
        string colgt6 = "WallThickness";
        string colgt7 = "Diameter";
        string colgt8 = "Grade";
        string colgt9 = "Coating";
        string colgt10 = "Manufacture";
        string colgt11 = "DoubleJointNo";


        string col_station = "STATION";
        string col_northing = "NORTHING";
        string col_easting = "EASTING";
        string col_elevation = "ELEVATION";
        string col_description = "DESCRIPTION";
        string col_type = "TYPE";
        string col_7 = "COLUMN 7";
        string col_8 = "COLUMN 8";
        string col_9 = "COLUMN 9";
        string col_10 = "COLUMN 10";
        string col_11 = "COLUMN 11";
        string col_12 = "COLUMN 12";
        string col_13 = "COLUMN 13";
        string col_14 = "COLUMN 14";
        string col_15 = "COLUMN 15";
        string col_16 = "COLUMN 16";
        string col_17 = "COLUMN 17";





        System.Data.DataTable dt_display;


        Form inputbox1 = null;
        private bool clickdragdown1;
        private System.Drawing.Point lastLocation1;


        public System.Data.DataTable dt_dismissed_errors = null;
        public string dismiss_errors_tab = "Dsd errors wm";
        Microsoft.Office.Interop.Excel.Worksheet W3 = null;

        string legend_cover_column = "XXXX";

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_pipe_rep);
            lista_butoane.Add(button_load_weld_map);
            lista_butoane.Add(button_wm_l);
            lista_butoane.Add(button_wm_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_create_weldmap);
            lista_butoane.Add(button_export_errors_to_xl);
            lista_butoane.Add(button_gen_wmR2);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_pipe_rep);
            lista_butoane.Add(button_load_weld_map);
            lista_butoane.Add(button_wm_l);
            lista_butoane.Add(button_wm_nl);
            lista_butoane.Add(button_refresh_ws1);
            lista_butoane.Add(button_create_weldmap);
            lista_butoane.Add(button_export_errors_to_xl);
            lista_butoane.Add(button_gen_wmR2);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Wgen_weldmap()
        {
            InitializeComponent();
            var toolStripMenuItem1 = new ToolStripMenuItem { Text = "Go to error" };
            toolStripMenuItem1.Click += go_to_excel_point;

            var toolStripMenuItem2 = new ToolStripMenuItem { Text = "Zoom to point in AutoCAD" };
            toolStripMenuItem2.Click += zoom_to_point_in_acad;

            ContextMenuStrip_go_to_error = new ContextMenuStrip();
            ContextMenuStrip_go_to_error.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2 });

        }









        private void make_first_line_invisible()
        {


            dt_display = new System.Data.DataTable();
            dt_display.Columns.Add("Point", typeof(string));
            dt_display.Columns.Add("Value", typeof(string));
            dt_display.Columns.Add("Excel", typeof(string));
            dt_display.Columns.Add("Error", typeof(string));
            dataGridView_error_weld_map.DataSource = dt_display;
            dataGridView_error_weld_map.Columns[0].Width = 75;
            dataGridView_error_weld_map.Columns[1].Width = 75;
            dataGridView_error_weld_map.Columns[2].Width = 50;
            dataGridView_error_weld_map.Columns[3].Width = 300;
            dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView_error_weld_map.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            dataGridView_error_weld_map.DefaultCellStyle.ForeColor = Color.White;
            dataGridView_error_weld_map.EnableHeadersVisualStyles = false;
        }

        private void button_refresh_ws1_Click(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            if (comboBox_ws1.Items.Count > 0)
            {
                for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
                {
                    if (comboBox_ws1.Items[i].ToString().ToUpper().Contains("WELD_MAP") == true)
                    {
                        comboBox_ws1.SelectedIndex = i;
                        i = comboBox_ws1.Items.Count;
                    }
                }
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_draw_pipes_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            if (Wgen_main_form.dt_weld_map == null || Wgen_main_form.dt_weld_map.Rows.Count == 0)
            {
                return;
            }

            Wgen_main_form.tpage_weldmap.Hide();
            Wgen_main_form.tpage_blank.Show();
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_blank.get_label_wait_visible(true);

            Wgen_main_form.tpage_blank.Refresh();

            set_enable_false();
            System.Data.DataTable dt1 = Functions.Creaza_weldmap_datatable_structure();
            System.Data.DataTable dt2 = Functions.Creaza_all_points_datatable_structure();
            int i = 0;

            for (i = 0; i < Wgen_main_form.dt_weld_map.Rows.Count; ++i)
            {
                if (Wgen_main_form.dt_weld_map.Rows[i][col_y] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col_x] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col_z] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col_feat_code] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col_sta_lin] != DBNull.Value &&
                    (Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToLower().Replace(" ", "") == "weld" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToLower().Replace(" ", "") == "wld" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToLower().Replace(" ", "") == "bend" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToLower().Replace(" ", "") == "elbow" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToLower().Replace(" ", "") == "bore_face" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_feat_code]).ToUpper().Replace(" ", "") == "LOOSE_END") &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_y])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_x])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_z])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_sta_lin]).Replace("+", "")) == true)
                {
                    dt1.ImportRow(Wgen_main_form.dt_weld_map.Rows[i]);
                }
            }

            #region import all points
            if (Wgen_main_form.dt_all_points != null && Wgen_main_form.dt_all_points.Rows.Count > 0)
            {
                for (i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                {
                    if (Wgen_main_form.dt_all_points.Rows[i][colpt1] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt2] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt3] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt4] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt5] != DBNull.Value &&
                         Wgen_main_form.dt_all_points.Rows[i][colpt6] != DBNull.Value &&
                            Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2])) == true &&
                            Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3])) == true &&
                            Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4])) == true
                         )
                    {
                        dt2.ImportRow(Wgen_main_form.dt_all_points.Rows[i]);
                    }
                }
            }
            #endregion

            if (dt1.Rows.Count == 0)
            {
                set_enable_true();
                Wgen_main_form.tpage_weldmap.Show();
                Wgen_main_form.tpage_blank.Hide();
                Wgen_main_form.tpage_pipe_manifest.Hide();
                Wgen_main_form.tpage_pipe_tally.Hide();
                Wgen_main_form.tpage_allpts.Hide();
                Wgen_main_form.tpage_build_pipe_tally.Hide();
                Wgen_main_form.tpage_duplicates.Hide();
                Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                return;
            }

            try
            {
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                        #region OBJECT DATA
                        Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                        Functions.Create_weldmap_od_table();
                        #endregion

                        List<string> lista_wt = new List<string>();
                        List<int> lista_cid = new List<int>();
                        int cid = 1;

                        for (i = 0; i < dt1.Rows.Count; ++i)
                        {
                            string fc1 = Convert.ToString(dt1.Rows[i][col_feat_code]).Replace(" ", "");
                            string fc2 = "";
                            double pipe_length = 0;
                            int last_i = i;

                            #region weld or loose end
                            if (fc1.ToLower() == "weld" || fc1.ToLower() == "loose_end" || fc1.ToLower() == "wld" || fc1.ToLower() == "le")
                            {
                                #region OBJECT DATA
                                List<object> Lista_val = new List<object>();
                                List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();

                                string STA_END = "";
                                string PNT_END = "";
                                string Xray2 = "";

                                #endregion

                                double x1 = Convert.ToDouble(dt1.Rows[i][col_x]);
                                double y1 = Convert.ToDouble(dt1.Rows[i][col_y]);
                                double z1 = Convert.ToDouble(dt1.Rows[i][col_z]);

                                double x2 = 0;
                                double y2 = 0;
                                double z2 = 0;

                                string wt_coating = "wt";
                                if (dt1.Rows[i][col_wall_ahd] != DBNull.Value)
                                {
                                    wt_coating = Convert.ToString(dt1.Rows[i][col_wall_ahd]);
                                }

                                if (dt1.Rows[i][col_coat_ahd] != DBNull.Value)
                                {
                                    wt_coating = wt_coating + "_" + Convert.ToString(dt1.Rows[i][col_coat_ahd]);
                                }

                                #region OBJECT DATA
                                string MMID = "";
                                if (dt1.Rows[i][col_mm_ahd] != DBNull.Value) MMID = Convert.ToString(dt1.Rows[i][col_mm_ahd]);

                                string PIPEID = "";
                                if (dt1.Rows[i][col_pipe_ahd] != DBNull.Value) PIPEID = Convert.ToString(dt1.Rows[i][col_pipe_ahd]);

                                string HEAT = "";
                                if (dt1.Rows[i][col_heat_ahd] != DBNull.Value) HEAT = Convert.ToString(dt1.Rows[i][col_heat_ahd]);

                                string Xray1 = "";
                                if (dt1.Rows[i][col_descr] != DBNull.Value) Xray1 = Convert.ToString(dt1.Rows[i][col_descr]);

                                string COATING = "";
                                if (dt1.Rows[i][col_coat_ahd] != DBNull.Value) COATING = Convert.ToString(dt1.Rows[i][col_coat_ahd]);

                                string STA_START = "";
                                if (dt1.Rows[i][col_sta_lin] != DBNull.Value)
                                {
                                    STA_START = Convert.ToString(dt1.Rows[i][col_sta_lin]);
                                    if (Functions.IsNumeric(STA_START.Replace("+", "")) == true)
                                    {
                                        STA_START = Functions.Get_chainage_from_double(Convert.ToDouble(STA_START.Replace("+", "")), "f", 2);
                                    }
                                }

                                string WALL = "";
                                if (dt1.Rows[i][col_wall_ahd] != DBNull.Value) WALL = Convert.ToString(dt1.Rows[i][col_wall_ahd]);

                                string PNT_START = "";
                                if (dt1.Rows[i][col_pnt] != DBNull.Value) PNT_START = Convert.ToString(dt1.Rows[i][col_pnt]);

                                int no_bore_face = 0;
                                PolylineVertex3d[] Vertex_new_bend = new PolylineVertex3d[0];
                                if (i < dt1.Rows.Count - 1)
                                {
                                    #region j
                                    for (int j = i + 1; j < dt1.Rows.Count; ++j)
                                    {
                                        fc2 = Convert.ToString(dt1.Rows[j][col_feat_code]).Replace(" ", "");

                                        if (fc2.ToLower() == "bend" || fc2.ToLower() == "elbow")
                                        {

                                            double xb = Convert.ToDouble(dt1.Rows[j][col_x]);
                                            double yb = Convert.ToDouble(dt1.Rows[j][col_y]);
                                            double zb = Convert.ToDouble(dt1.Rows[j][col_z]);

                                            Array.Resize(ref Vertex_new_bend, Vertex_new_bend.Length + 1);
                                            Vertex_new_bend[Vertex_new_bend.Length - 1] = new PolylineVertex3d(new Point3d(xb, yb, zb));

                                            #region Point_block

                                            System.Collections.Specialized.StringCollection Col_name1 = new System.Collections.Specialized.StringCollection();
                                            System.Collections.Specialized.StringCollection Col_value1 = new System.Collections.Specialized.StringCollection();

                                            string PNT_BEND = "";
                                            if (dt1.Rows[j][col_pnt] != DBNull.Value) PNT_BEND = Convert.ToString(dt1.Rows[j][col_pnt]);

                                            Col_name1.Add("PTNO");
                                            Col_value1.Add(PNT_BEND);

                                            Col_name1.Add("FEATURE_CODE");
                                            if (dt1.Rows[j][col_feat_code] != DBNull.Value) Col_value1.Add(Convert.ToString(dt1.Rows[j][col_feat_code]));

                                            string DESC1 = "";
                                            if (dt1.Rows[j][col_descr] != DBNull.Value) DESC1 = Convert.ToString(dt1.Rows[j][col_descr]);
                                            Col_name1.Add("DESCRIPTION");
                                            Col_value1.Add(DESC1);

                                            string STA_b = "";
                                            if (dt1.Rows[j][col_sta_lin] != DBNull.Value)
                                            {
                                                STA_b = Convert.ToString(dt1.Rows[j][col_sta_lin]);
                                                if (Functions.IsNumeric(STA_b.Replace("+", "")) == true)
                                                {
                                                    STA_b = Functions.Get_chainage_from_double(Convert.ToDouble(STA_b.Replace("+", "")), "f", 2);
                                                }
                                            }

                                            Col_name1.Add("STATION");
                                            Col_value1.Add(STA_b);

                                            BlockReference block44 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, Wgen_main_form.WGEN_folder + "4.dwg", "4", new Point3d(xb, yb, zb), 1,
                                                                                                                                0, "0", Col_name1, Col_value1);

                                            if (block44 == null)
                                            {
                                                MessageBox.Show("the block 4 was not found");
                                                set_enable_true();
                                                Editor1.SetImpliedSelection(Empty_array);
                                                Editor1.WriteMessage("\nCommand:");
                                                Wgen_main_form.tpage_weldmap.Show();
                                                Wgen_main_form.tpage_blank.Hide();
                                                Wgen_main_form.tpage_pipe_manifest.Hide();
                                                Wgen_main_form.tpage_pipe_tally.Hide();
                                                Wgen_main_form.tpage_allpts.Hide();
                                                Wgen_main_form.tpage_build_pipe_tally.Hide();
                                                Wgen_main_form.tpage_duplicates.Hide();
                                                Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                                                return;
                                            }
                                            #endregion

                                            #region OBJECT DATA
                                            if (dt1.Rows[j][col_pnt] != DBNull.Value)
                                            {
                                                if (PNT_END == "")
                                                {
                                                    PNT_END = Convert.ToString(dt1.Rows[j][col_pnt]);
                                                }
                                                else
                                                {
                                                    PNT_END = PNT_END + ", " + Convert.ToString(dt1.Rows[j][col_pnt]);
                                                }

                                            }

                                            if ((dt1.Rows[j][col_descr]) != DBNull.Value)
                                            {
                                                Xray2 = Convert.ToString(dt1.Rows[j][col_descr]);
                                            }
                                            #endregion

                                        }
                                        else if (fc2.ToLower() == "weld" || fc2.ToLower() == "loose_end" || fc2.ToLower() == "wld" || fc2.ToLower() == "le")
                                        {
                                            x2 = Convert.ToDouble(dt1.Rows[j][col_x]);
                                            y2 = Convert.ToDouble(dt1.Rows[j][col_y]);
                                            z2 = Convert.ToDouble(dt1.Rows[j][col_z]);

                                            #region OBJECT DATA
                                            if (dt1.Rows[j][col_pnt] != DBNull.Value)
                                            {
                                                if (PNT_END != "")
                                                {
                                                    PNT_END = PNT_END + ", " + Convert.ToString(dt1.Rows[j][col_pnt]);
                                                }
                                                else
                                                {
                                                    PNT_END = Convert.ToString(dt1.Rows[j][col_pnt]);
                                                }
                                            }

                                            if (dt1.Rows[j][col_sta_lin] != DBNull.Value)
                                            {
                                                STA_END = Convert.ToString(dt1.Rows[j][col_sta_lin]);
                                                if (Functions.IsNumeric(STA_END.Replace("+", "")) == true)
                                                {
                                                    STA_END = Functions.Get_chainage_from_double(Convert.ToDouble(STA_END.Replace("+", "")), "f", 2);
                                                }
                                            }

                                            if ((dt1.Rows[j][col_descr]) != DBNull.Value)
                                            {
                                                Xray2 = Convert.ToString(dt1.Rows[j][col_descr]);
                                            }

                                            if (fc2.ToLower() == "loose_end")
                                            {
                                                if ((fc1.ToLower() != "loose_end"))
                                                {
                                                    last_i = j;
                                                }
                                                if (fc1.ToLower() == "loose_end" && no_bore_face == 2)
                                                {
                                                    last_i = j - 1;
                                                }
                                            }

                                            #endregion
                                            j = dt1.Rows.Count;
                                        }
                                        else if (fc2.ToLower() == "bore_face")
                                        {
                                            ++no_bore_face;
                                        }

                                    }
                                    #endregion
                                }

                                Lista_val.Add(MMID);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(PIPEID);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(HEAT);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(Xray1);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(Xray2);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(COATING);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(WALL);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(STA_START);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(STA_END);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(PNT_START);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                Lista_val.Add(PNT_END);
                                Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);



                                #endregion

                                #region COLOR INDEX
                                if (lista_wt.Contains(wt_coating) == false)
                                {
                                    lista_wt.Add(wt_coating);
                                    lista_cid.Add(cid);

                                    if (cid == 1)
                                    {
                                        cid = 2;
                                    }
                                    else if (cid == 2)
                                    {
                                        cid = 5;
                                    }
                                    else if (cid == 5)
                                    {
                                        cid = 6;
                                    }
                                    else if (cid == 6)
                                    {
                                        cid = 84;
                                    }
                                    else if (cid == 84)
                                    {
                                        cid = 172;
                                    }
                                    else if (cid == 172)
                                    {
                                        cid = 53;
                                    }
                                    else if (cid == 53)
                                    {
                                        cid = 1;
                                    }

                                }
                                #endregion

                                System.Collections.Specialized.StringCollection Col_name = new System.Collections.Specialized.StringCollection();
                                System.Collections.Specialized.StringCollection Col_value = new System.Collections.Specialized.StringCollection();

                                Col_name.Add("PIPEID");
                                Col_value.Add(PIPEID);
                                Col_name.Add("HEATID");
                                Col_value.Add(HEAT);


                                #region point block
                                Col_name.Add("PTNO");
                                Col_value.Add(PNT_START);



                                Col_name.Add("FEATURE_CODE");
                                Col_value.Add(fc1);

                                Col_name.Add("STATION");
                                Col_value.Add(STA_START);

                                if (fc1 == "WELD" || fc1 == "WLD")
                                {
                                    Col_name.Add("XRAY");
                                    Col_value.Add(Xray1);
                                    if (dt1.Rows[i][col_mm_bk] != DBNull.Value)
                                    {
                                        string MM_BACK = Convert.ToString(dt1.Rows[i][col_mm_bk]);
                                        Col_name.Add("MM_BACK");
                                        Col_value.Add(MM_BACK);
                                    }

                                    if (dt1.Rows[i][col_mm_ahd] != DBNull.Value)
                                    {
                                        string MM_AHEAD = Convert.ToString(dt1.Rows[i][col_mm_ahd]);
                                        Col_name.Add("MM_AHEAD");
                                        Col_value.Add(MM_AHEAD);
                                    }
                                }
                                else
                                {



                                    Col_name.Add("DESCRIPTION");
                                    Col_value.Add(Xray1);



                                }



                                #endregion


                                if (last_i > i)
                                {
                                    for (int k = i + 1; k <= last_i; ++k)
                                    {
                                        #region point block

                                        System.Collections.Specialized.StringCollection Col_namek = new System.Collections.Specialized.StringCollection();
                                        System.Collections.Specialized.StringCollection Col_valuek = new System.Collections.Specialized.StringCollection();

                                        string pnt1 = "";
                                        if (dt1.Rows[k][col_pnt] != DBNull.Value) pnt1 = Convert.ToString(dt1.Rows[k][col_pnt]);
                                        Col_namek.Add("PTNO");
                                        Col_valuek.Add(pnt1);


                                        string fc = "";
                                        if (dt1.Rows[k][col_feat_code] != DBNull.Value) fc = Convert.ToString(dt1.Rows[k][col_feat_code]).Replace(" ", "");
                                        Col_namek.Add("FEATURE_CODE");
                                        Col_valuek.Add(fc);

                                        string sta1 = "";
                                        if (dt1.Rows[k][col_sta_lin] != DBNull.Value)
                                        {
                                            sta1 = Convert.ToString(dt1.Rows[k][col_sta_lin]);
                                            if (Functions.IsNumeric(sta1.Replace("+", "")) == true)
                                            {
                                                sta1 = Functions.Get_chainage_from_double(Convert.ToDouble(sta1.Replace("+", "")), "f", 2);
                                            }
                                        }
                                        Col_namek.Add("STATION");
                                        Col_valuek.Add(sta1);

                                        string descr = "";
                                        if (dt1.Rows[k][col_descr] != DBNull.Value) descr = Convert.ToString(dt1.Rows[k][col_descr]);
                                        Col_namek.Add("DESCRIPTION");
                                        Col_valuek.Add(descr);

                                        double xk = Convert.ToDouble(dt1.Rows[k][col_x]);
                                        double yk = Convert.ToDouble(dt1.Rows[k][col_y]);
                                        double zk = Convert.ToDouble(dt1.Rows[k][col_z]);



                                        BlockReference block5 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, Wgen_main_form.WGEN_folder + "4.dwg", "4", new Point3d(xk, yk, zk), 1,
                                                                                                                            0, "0", Col_namek, Col_valuek);
                                        if (block5 == null)
                                        {
                                            MessageBox.Show("the block 4 was not found");
                                            set_enable_true();
                                            Editor1.SetImpliedSelection(Empty_array);
                                            Editor1.WriteMessage("\nCommand:");
                                            Wgen_main_form.tpage_weldmap.Show();
                                            Wgen_main_form.tpage_blank.Hide();
                                            Wgen_main_form.tpage_pipe_manifest.Hide();
                                            Wgen_main_form.tpage_pipe_tally.Hide();
                                            Wgen_main_form.tpage_allpts.Hide();
                                            Wgen_main_form.tpage_build_pipe_tally.Hide();
                                            Wgen_main_form.tpage_duplicates.Hide();
                                            Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                                            return;
                                        }
                                        #endregion
                                    }
                                }




                                if (i < dt1.Rows.Count - 1)
                                {

                                    Polyline3d pipe1 = new Polyline3d();

                                    int color1 = lista_cid[lista_wt.IndexOf(wt_coating)];
                                    pipe1.ColorIndex = color1;
                                    BTrecord.AppendEntity(pipe1);
                                    Trans1.AddNewlyCreatedDBObject(pipe1, true);

                                    PolylineVertex3d Vertex_new1 = new PolylineVertex3d(new Point3d(x1, y1, z1));
                                    pipe1.AppendVertex(Vertex_new1);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new1, true);

                                    if (Vertex_new_bend.Length > 0)
                                    {
                                        for (int k = 0; k < Vertex_new_bend.Length; ++k)
                                        {
                                            pipe1.AppendVertex(Vertex_new_bend[k]);
                                            Trans1.AddNewlyCreatedDBObject(Vertex_new_bend[k], true);
                                        }
                                    }

                                    PolylineVertex3d Vertex_new2 = new PolylineVertex3d(new Point3d(x2, y2, z2));
                                    pipe1.AppendVertex(Vertex_new2);
                                    Trans1.AddNewlyCreatedDBObject(Vertex_new2, true);
                                    Functions.Populate_object_data_table_from_objectid(Tables1, pipe1.ObjectId, "WGEN_PIPE_REP", Lista_val, Lista_type);

                                    pipe_length = pipe1.Length;

                                    double len2 = 0;
                                    if (Functions.IsNumeric(STA_START.Replace("+", "")) == true && Functions.IsNumeric(STA_END.Replace("+", "")) == true)
                                    {
                                        len2 = Convert.ToDouble(STA_END.Replace("+", "")) - Convert.ToDouble(STA_START.Replace("+", ""));
                                    }

                                    string lenstr = "";

                                    if (Math.Abs(pipe_length - len2) > 0.5)
                                    {
                                        lenstr = Functions.Get_String_Rounded(pipe_length, 1) + "(" + Functions.Get_String_Rounded(len2, 1) + ")";

                                    }
                                    else
                                    {
                                        lenstr = Functions.Get_String_Rounded(pipe_length, 1);
                                    }
                                    Col_name.Add("LENGTH");
                                    Col_value.Add(lenstr);






                                    double dist1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                                    string blockname1 = "11";
                                    if (dist1 < 20) blockname1 = "22";
                                    if (dist1 < 3) blockname1 = "33";

                                    double rot1 = Functions.GET_Bearing_rad(x1, y1, x2, y2);
                                    Point3d inspt = new Point3d(x1, y1, z1);

                                    if (rot1 > Math.PI / 2 && rot1 < 3 * Math.PI / 2)
                                    {
                                        rot1 = rot1 + Math.PI;
                                        inspt = new Point3d(x2, y2, z2);
                                    }


                                    BlockReference block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, Wgen_main_form.WGEN_folder + blockname1 + ".dwg", blockname1, inspt, 1,
                                                       rot1, "0", Col_name, Col_value);
                                    if (block1 == null)
                                    {
                                        MessageBox.Show("the block " + blockname1 + " was not found");
                                        set_enable_true();
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        Wgen_main_form.tpage_weldmap.Show();
                                        Wgen_main_form.tpage_blank.Hide();
                                        Wgen_main_form.tpage_pipe_manifest.Hide();
                                        Wgen_main_form.tpage_pipe_tally.Hide();
                                        Wgen_main_form.tpage_allpts.Hide();
                                        Wgen_main_form.tpage_build_pipe_tally.Hide();
                                        Wgen_main_form.tpage_duplicates.Hide();
                                        Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                                        return;
                                    }

                                    Functions.Stretch_block(block1, "Distance1", dist1);
                                    block1.ColorIndex = lista_cid[lista_wt.IndexOf(wt_coating)];

                                }

                                #region point block
                                BlockReference block4 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, Wgen_main_form.WGEN_folder + "4.dwg", "4", new Point3d(x1, y1, z1), 1,
                                                                                                                    0, "0", Col_name, Col_value);
                                if (block4 == null)
                                {
                                    MessageBox.Show("the block 4 was not found");
                                    set_enable_true();
                                    Editor1.SetImpliedSelection(Empty_array);
                                    Editor1.WriteMessage("\nCommand:");
                                    Wgen_main_form.tpage_weldmap.Show();
                                    Wgen_main_form.tpage_blank.Hide();
                                    Wgen_main_form.tpage_pipe_manifest.Hide();
                                    Wgen_main_form.tpage_pipe_tally.Hide();
                                    Wgen_main_form.tpage_allpts.Hide();
                                    Wgen_main_form.tpage_build_pipe_tally.Hide();
                                    Wgen_main_form.tpage_duplicates.Hide();
                                    Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                                    return;
                                }
                                #endregion

                            }


                            #endregion

                            i = last_i;
                        }

                        #region all points....
                        if (dt2.Rows.Count > 0)
                        {
                            for (i = 0; i < dt2.Rows.Count - 1; ++i)
                            {
                                string ft_code = Convert.ToString(dt2.Rows[i][colpt5]);
                                if ((ft_code.ToUpper() != "WELD" || ft_code.ToUpper() != "WLD") && ft_code.ToUpper() != "BEND" && ft_code.ToUpper() != "ELBOW")
                                {
                                    System.Collections.Specialized.StringCollection Col_name = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Col_value = new System.Collections.Specialized.StringCollection();

                                    double yng1 = Convert.ToDouble(dt2.Rows[i][colpt2]);
                                    double xng1 = Convert.ToDouble(dt2.Rows[i][colpt3]);
                                    double zng1 = Convert.ToDouble(dt2.Rows[i][colpt4]);
                                    string ptno = Convert.ToString(dt2.Rows[i][colpt1]);

                                    string st_string = "";
                                    if (Functions.IsNumeric(Convert.ToString(dt2.Rows[i][colpt6]).Replace("+", "")) == true)
                                    {
                                        st_string = Functions.Get_chainage_from_double(Convert.ToDouble(Convert.ToString(dt2.Rows[i][colpt6]).Replace("+", "")), "f", 2);
                                    }
                                    Col_name.Add("PTNO");
                                    Col_value.Add(ptno);
                                    Col_name.Add("LINE1");
                                    Col_value.Add(st_string);
                                    Col_name.Add("LINE2");
                                    Col_value.Add(ft_code);

                                    string val8 = "";
                                    if (dt2.Rows[i][colpt8] != DBNull.Value)
                                    {
                                        val8 = Convert.ToString(dt2.Rows[i][colpt8]);
                                    }
                                    string val9 = "";
                                    if (dt2.Rows[i][colpt9] != DBNull.Value)
                                    {
                                        val9 = Convert.ToString(dt2.Rows[i][colpt9]);
                                    }
                                    string val10 = "";
                                    if (dt2.Rows[i][colpt10] != DBNull.Value)
                                    {
                                        val10 = Convert.ToString(dt2.Rows[i][colpt10]);
                                    }
                                    string val11 = "";
                                    if (dt2.Rows[i][colpt11] != DBNull.Value)
                                    {
                                        val11 = Convert.ToString(dt2.Rows[i][colpt11]);
                                    }
                                    string val12 = "";
                                    if (dt2.Rows[i][colpt12] != DBNull.Value)
                                    {
                                        val12 = Convert.ToString(dt2.Rows[i][colpt12]);
                                    }
                                    string val13 = "";
                                    if (dt2.Rows[i][colpt13] != DBNull.Value)
                                    {
                                        val13 = Convert.ToString(dt2.Rows[i][colpt13]);
                                    }
                                    string val14 = "";
                                    if (dt2.Rows[i][colpt14] != DBNull.Value)
                                    {
                                        val14 = Convert.ToString(dt2.Rows[i][colpt14]);
                                    }
                                    string val15 = "";
                                    if (dt2.Rows[i][colpt15] != DBNull.Value)
                                    {
                                        val15 = Convert.ToString(dt2.Rows[i][colpt15]);
                                    }
                                    string val16 = "";
                                    if (dt2.Rows[i][colpt16] != DBNull.Value)
                                    {
                                        val16 = Convert.ToString(dt2.Rows[i][colpt16]);
                                    }


                                    Col_name.Add("LINE3");
                                    Col_value.Add(val8);
                                    Col_name.Add("LINE4");
                                    Col_value.Add(val9);
                                    Col_name.Add("LINE5");
                                    Col_value.Add(val10);
                                    Col_name.Add("LINE6");
                                    Col_value.Add(val11);
                                    Col_name.Add("LINE7");
                                    Col_value.Add(val12);
                                    Col_name.Add("LINE8");
                                    Col_value.Add(val13);
                                    Col_name.Add("LINE9");
                                    Col_value.Add(val14);
                                    Col_name.Add("LINE10");
                                    Col_value.Add(val15);
                                    Col_name.Add("LINE11");
                                    Col_value.Add(val16);


                                    BlockReference block5 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, Wgen_main_form.WGEN_folder + "5.dwg", "5", new Point3d(xng1, yng1, zng1), 1,
                                                    0, "0", Col_name, Col_value);
                                    block5.ColorIndex = 3;

                                    if (block5 == null)
                                    {
                                        MessageBox.Show("the block 5 was not found");
                                        set_enable_true();
                                        Editor1.SetImpliedSelection(Empty_array);
                                        Editor1.WriteMessage("\nCommand:");
                                        Wgen_main_form.tpage_weldmap.Show();
                                        Wgen_main_form.tpage_blank.Hide();
                                        Wgen_main_form.tpage_pipe_manifest.Hide();
                                        Wgen_main_form.tpage_pipe_tally.Hide();
                                        Wgen_main_form.tpage_allpts.Hide();
                                        Wgen_main_form.tpage_build_pipe_tally.Hide();
                                        Wgen_main_form.tpage_duplicates.Hide();
                                        Wgen_main_form.tpage_blank.get_label_wait_visible(false);
                                        return;
                                    }
                                }
                            }
                        }
                        #endregion

                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(i.ToString() + ":\r\n" + ex.Message);
            }
            set_enable_true();
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            Wgen_main_form.tpage_weldmap.Show();
            Wgen_main_form.tpage_blank.Hide();
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_blank.get_label_wait_visible(false);
        }



        /// <summary>
        ///  adds a column station into the all points table
        /// </summary>
        /// <param name="cl_poly">polyline that came from alignment</param>


        private double Station_equation_of(double Station_measured)
        {
            double Valoare = 0;
            double Valoare_de_returnat = Station_measured + Valoare;

            if (dt_st_eq != null)
            {
                if (dt_st_eq.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_st_eq.Rows.Count; ++i)
                    {
                        if (dt_st_eq.Rows[i]["Back"] != DBNull.Value && dt_st_eq.Rows[i]["Ahead"] != DBNull.Value)
                        {
                            double Station_back = Convert.ToDouble(dt_st_eq.Rows[i]["Back"]);
                            double Station_ahead = Convert.ToDouble(dt_st_eq.Rows[i]["Ahead"]);
                            if (Station_measured + Valoare < Station_back)
                            {
                                return Station_measured + Valoare;
                            }
                            else
                            {
                                Valoare = Valoare + Station_ahead - Station_back;
                                Valoare_de_returnat = Station_measured + Valoare;
                            }
                        }
                    }
                }
            }
            return Valoare_de_returnat;
        }

        public DialogResult InputBox(string title, string promptText)
        {
            int xright = 200;
            int height1 = 120;

            inputbox1 = new Form();
            System.Windows.Forms.Label label1 = new System.Windows.Forms.Label();
            System.Windows.Forms.Label label2 = new System.Windows.Forms.Label();
            System.Windows.Forms.Label title1 = new System.Windows.Forms.Label();
            textbox1 = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOK = new System.Windows.Forms.Button();
            System.Windows.Forms.Panel Panel1 = new System.Windows.Forms.Panel();
            System.Windows.Forms.Button buttonExit = new System.Windows.Forms.Button();

            Color PanelRgbColor = new Color();
            PanelRgbColor = Color.FromArgb(0, 122, 204);
            Panel1.BackColor = PanelRgbColor;
            Panel1.SetBounds(0, 0, 396, 25);
            Panel1.BorderStyle = BorderStyle.None;
            Panel1.Controls.AddRange(new Control[] { buttonExit, title1 });
            Panel1.MouseDown += new MouseEventHandler(clickmove_MouseDown1);
            Panel1.MouseMove += new MouseEventHandler(clickmove_MouseMove1);
            Panel1.MouseUp += new MouseEventHandler(clickmove_MouseUp1);

            buttonExit.BackgroundImage = ((System.Drawing.Image)(Properties.Resources.close));
            buttonExit.BackgroundImageLayout = ImageLayout.Stretch;
            buttonExit.SetBounds(xright + 10 - 20, 3, 20, 20);
            buttonExit.Click += new EventHandler(button_close_Click);

            title1.Text = title;
            System.Drawing.Font Font2 = new System.Drawing.Font("Arial", 12f, FontStyle.Bold);
            title1.Font = Font2;
            title1.ForeColor = Color.White;
            title1.SetBounds(2, 3, 100, 13);
            title1.AutoSize = true;
            title1.MouseDown += new MouseEventHandler(clickmove_MouseDown1);
            title1.MouseMove += new MouseEventHandler(clickmove_MouseMove1);
            title1.MouseUp += new MouseEventHandler(clickmove_MouseUp1);

            label1.Text = promptText;
            System.Drawing.Font Font1 = new System.Drawing.Font("Arial", 12f, FontStyle.Bold);
            label1.Font = Font1;
            label1.ForeColor = Color.White;
            label1.SetBounds(15, 35, 372, 13);
            label1.AutoSize = true;

            label2.Text = "Length check tolerance:";
            System.Drawing.Font Font3 = new System.Drawing.Font("Arial", 9f, FontStyle.Bold);
            label2.Font = Font3;
            label2.ForeColor = Color.White;
            label2.SetBounds(15, 60, 372, 13);
            label2.AutoSize = true;


            buttonOK.Text = "OK";
            buttonOK.DialogResult = DialogResult.OK;
            buttonOK.ForeColor = Color.White;
            buttonOK.SetBounds(xright - 75, height1 - 10 - 25, 75, 25);
            buttonOK.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;


            textbox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(51)))), ((int)(((byte)(55)))));
            textbox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            textbox1.ForeColor = System.Drawing.Color.White;
            textbox1.Location = new System.Drawing.Point(xright - 35, 58);
            textbox1.Size = new System.Drawing.Size(20, 20);
            textbox1.Text = "1";
            textbox1.Name = "Txt_tolerance";


            Color color1 = new Color();
            color1 = Color.FromArgb(37, 37, 38);
            inputbox1.BackColor = color1;
            inputbox1.Text = title;
            inputbox1.MinimumSize = new Size(xright + 10, height1);
            inputbox1.ClientSize = new Size(xright + 10, height1);


            inputbox1.Controls.AddRange(new Control[] { label1, label2, textbox1, buttonOK, Panel1 });

            inputbox1.FormBorderStyle = FormBorderStyle.None;
            inputbox1.StartPosition = FormStartPosition.CenterScreen;
            inputbox1.MinimizeBox = false;
            inputbox1.MaximizeBox = false;
            inputbox1.AcceptButton = buttonOK;



            DialogResult dialogResult = inputbox1.ShowDialog();

            if (Functions.IsNumeric(textbox1.Text) == true)
            {
                length_check_tolerance = Convert.ToDouble(textbox1.Text);
            }

            else
            {
                length_check_tolerance = 1;
            }

            return dialogResult;

        }
        private void button_close_Click(object sender, EventArgs e)
        {

            length_check_tolerance = 1;
            inputbox1.Close();

        }

        private void clickmove_MouseDown1(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown1 = true;
            lastLocation1 = e.Location;
        }

        private void clickmove_MouseMove1(object sender, MouseEventArgs e)
        {
            if (clickdragdown1)
            {
                inputbox1.Location = new System.Drawing.Point(
                  (inputbox1.Location.X - lastLocation1.X) + e.X, (inputbox1.Location.Y - lastLocation1.Y) + e.Y);

                inputbox1.Update();
            }
        }

        private void clickmove_MouseUp1(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            clickdragdown1 = false;
        }

        private void button_create_weldmap_Click(object sender, EventArgs e)
        {



            if (Wgen_main_form.dt_all_points == null || Wgen_main_form.dt_all_points.Rows.Count == 0 || Wgen_main_form.dt_ground_tally == null || Wgen_main_form.dt_ground_tally.Rows.Count == 0)
            {
                return;
            }





            set_enable_false();
            int i = 0;
            dt_st_eq = null;
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        if (InputBox("WGEN", "Select the Centerline:") == DialogResult.OK)
                        {

                        }
                        else
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt5;
                        Prompt5 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt5.SetRejectMessage("\nSelect a polyline!");
                        Prompt5.AllowNone = true;
                        Prompt5.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt5.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Alignment), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt5);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {

                            MessageBox.Show("no centerline");

                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }

                        Polyline Poly_cl = null;
                        Poly_cl = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        bool is_from_alignment = false;
                        if (Poly_cl == null)
                        {
                            Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                            if (align1 != null)
                            {
                                Poly_cl = Trans1.GetObject(align1.GetPolyline(), OpenMode.ForRead) as Polyline;

                                if (Poly_cl == null)
                                {
                                    MessageBox.Show("error at the alignment object");
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double Start1 = align1.StartingStation;
                                is_from_alignment = true;
                                Autodesk.Civil.DatabaseServices.StationEquationCollection col_st_eq = align1.StationEquations;

                                if (col_st_eq.Count > 0 || Start1 != 0)
                                {
                                    dt_st_eq = new System.Data.DataTable();
                                    dt_st_eq.Columns.Add("Back", typeof(double));
                                    dt_st_eq.Columns.Add("Ahead", typeof(double));
                                    if (Start1 != 0)
                                    {
                                        dt_st_eq.Rows.Add();
                                        dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Back"] = 0;
                                        dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Ahead"] = Start1;
                                    }
                                    if (col_st_eq.Count > 0)
                                    {
                                        for (i = 0; i < col_st_eq.Count; ++i)
                                        {
                                            Autodesk.Civil.DatabaseServices.StationEquation eq1 = col_st_eq[i];
                                            dt_st_eq.Rows.Add();
                                            dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Back"] = eq1.StationBack;
                                            dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Ahead"] = eq1.StationAhead;
                                        }
                                    }

                                    dt_st_eq = Functions.Sort_data_table(dt_st_eq, "Back");
                                }



                            }
                        }

                        if (Poly_cl != null)
                        {
                            Add_station_to_all_points(Poly_cl);
                            if (is_from_alignment == true)
                            {
                                BTrecord.UpgradeOpen();
                                Poly_cl.UpgradeOpen();
                                Poly_cl.Erase();
                            }
                        }
                        else
                        {
                            MessageBox.Show("no centerline");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                set_enable_true();
                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                this.MdiParent.WindowState = FormWindowState.Normal;
                return;
            }

            this.MdiParent.WindowState = FormWindowState.Normal;

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");

            System.Data.DataTable dt_ng = new System.Data.DataTable();
            dt_ng.Columns.Add("ptno", typeof(string));
            dt_ng.Columns.Add("point", typeof(Point3d));

            System.Data.DataTable dt_colors = new System.Data.DataTable();
            dt_colors.Columns.Add("color", typeof(int));
            dt_colors.Columns.Add("colorindex", typeof(int));
            dt_colors.Columns.Add("themecolor", typeof(int));
            dt_colors.Columns.Add("tint", typeof(double));


            try
            {
                Wgen_main_form.dt_all_points = Functions.Sort_data_table(Wgen_main_form.dt_all_points, colpt_sta);
                Wgen_main_form.dt_all_points.TableName = "ALLPTS";
                Wgen_main_form.dt_ground_tally.TableName = "PIPETALLY";



                Wgen_main_form.dt_weld_map = Functions.Creaza_weldmap_datatable_structure();

                Wgen_main_form.tpage_weldmap.Hide();
                Wgen_main_form.tpage_blank.Show();
                Wgen_main_form.tpage_pipe_manifest.Hide();
                Wgen_main_form.tpage_pipe_tally.Hide();
                Wgen_main_form.tpage_allpts.Hide();
                Wgen_main_form.tpage_build_pipe_tally.Hide();
                Wgen_main_form.tpage_duplicates.Hide();
                Wgen_main_form.tpage_blank.get_label_wait_visible(true);

                this.Refresh();


                string col_bend_type = colpt7;
                string col_bend_defl = colpt8;
                string col_bend_pos = colpt9;
                string col_bend_hor = colpt10;
                string col_bend_ver = colpt11;

                string col_elbow_type = colpt7;
                string col_elbow_defl = colpt8;
                string col_elbow_pos = colpt9;
                string col_elbow_hor = colpt10;
                string col_elbow_ver = colpt11;

                string col_mm_back = colpt9;
                string col_mm_ahead = colpt10;

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
                                                col_bend_type = colpt6;

                                            }
                                            if (check15.Contains("{G}") == true)
                                            {
                                                col_bend_type = colpt7;

                                            }
                                            if (check15.Contains("{H}") == true)
                                            {
                                                col_bend_type = colpt8;

                                            }
                                            if (check15.Contains("{I}") == true)
                                            {
                                                col_bend_type = colpt9;

                                            }
                                            if (check15.Contains("{J}") == true)
                                            {
                                                col_bend_type = colpt10;

                                            }
                                            if (check15.Contains("{K}") == true)
                                            {
                                                col_bend_type = colpt11;

                                            }
                                            if (check15.Contains("{L}") == true)
                                            {
                                                col_bend_type = colpt12;

                                            }
                                            if (check15.Contains("{M}") == true)
                                            {
                                                col_bend_type = colpt13;
                                            }
                                            if (check15.Contains("{N}") == true)
                                            {
                                                col_bend_type = colpt14;
                                            }
                                            if (check15.Contains("{O}") == true)
                                            {
                                                col_bend_type = colpt15;
                                            }
                                            if (check15.Contains("{P}") == true)
                                            {
                                                col_bend_type = colpt16;
                                            }

                                            if (check15.Contains("{Q}") == true)
                                            {
                                                col_bend_type = colpt17;
                                            }
                                            if (check15.Contains("{R}") == true)
                                            {
                                                col_bend_type = colpt18;
                                            }
                                            if (check15.Contains("{S}") == true)
                                            {
                                                col_bend_type = colpt19;
                                            }
                                            if (check15.Contains("{T}") == true)
                                            {
                                                col_bend_type = colpt20;
                                            }
                                            if (check15.Contains("{U}") == true)
                                            {
                                                col_bend_type = colpt21;
                                            }
                                            if (check15.Contains("{V}") == true)
                                            {
                                                col_bend_type = colpt22;
                                            }
                                            if (check15.Contains("{W}") == true)
                                            {
                                                col_bend_type = colpt23;
                                            }
                                            if (check15.Contains("{X}") == true)
                                            {
                                                col_bend_type = colpt24;
                                            }
                                            if (check15.Contains("{Y}") == true)
                                            {
                                                col_bend_type = colpt25;
                                            }
                                            if (check15.Contains("{Z}") == true)
                                            {
                                                col_bend_type = colpt26;
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
                                                col_bend_defl = colpt6;

                                            }
                                            if (check16.Contains("{G}") == true)
                                            {
                                                col_bend_defl = colpt7;

                                            }
                                            if (check16.Contains("{H}") == true)
                                            {
                                                col_bend_defl = colpt8;

                                            }
                                            if (check16.Contains("{I}") == true)
                                            {
                                                col_bend_defl = colpt9;

                                            }
                                            if (check16.Contains("{J}") == true)
                                            {
                                                col_bend_defl = colpt10;

                                            }
                                            if (check16.Contains("{K}") == true)
                                            {
                                                col_bend_defl = colpt11;

                                            }
                                            if (check16.Contains("{L}") == true)
                                            {
                                                col_bend_defl = colpt12;

                                            }
                                            if (check16.Contains("{M}") == true)
                                            {
                                                col_bend_defl = colpt13;

                                            }
                                            if (check16.Contains("{N}") == true)
                                            {
                                                col_bend_defl = colpt14;

                                            }
                                            if (check16.Contains("{O}") == true)
                                            {
                                                col_bend_defl = colpt15;

                                            }
                                            if (check16.Contains("{P}") == true)
                                            {
                                                col_bend_defl = colpt16;

                                            }
                                            if (check16.Contains("{Q}") == true)
                                            {
                                                col_bend_defl = colpt17;
                                            }
                                            if (check16.Contains("{R}") == true)
                                            {
                                                col_bend_defl = colpt18;
                                            }
                                            if (check16.Contains("{S}") == true)
                                            {
                                                col_bend_defl = colpt19;
                                            }
                                            if (check16.Contains("{T}") == true)
                                            {
                                                col_bend_defl = colpt20;
                                            }
                                            if (check16.Contains("{U}") == true)
                                            {
                                                col_bend_defl = colpt21;
                                            }
                                            if (check16.Contains("{V}") == true)
                                            {
                                                col_bend_defl = colpt22;
                                            }
                                            if (check16.Contains("{W}") == true)
                                            {
                                                col_bend_defl = colpt23;
                                            }
                                            if (check16.Contains("{X}") == true)
                                            {
                                                col_bend_defl = colpt24;
                                            }
                                            if (check16.Contains("{Y}") == true)
                                            {
                                                col_bend_defl = colpt25;
                                            }
                                            if (check16.Contains("{Z}") == true)
                                            {
                                                col_bend_defl = colpt26;
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
                                                col_bend_pos = colpt6;

                                            }
                                            if (check17.Contains("{G}") == true)
                                            {
                                                col_bend_pos = colpt7;

                                            }
                                            if (check17.Contains("{H}") == true)
                                            {
                                                col_bend_pos = colpt8;

                                            }
                                            if (check17.Contains("{I}") == true)
                                            {
                                                col_bend_pos = colpt9;

                                            }
                                            if (check17.Contains("{J}") == true)
                                            {
                                                col_bend_pos = colpt10;

                                            }
                                            if (check17.Contains("{K}") == true)
                                            {
                                                col_bend_pos = colpt11;

                                            }
                                            if (check17.Contains("{L}") == true)
                                            {
                                                col_bend_pos = colpt12;

                                            }
                                            if (check17.Contains("{M}") == true)
                                            {
                                                col_bend_pos = colpt13;

                                            }
                                            if (check17.Contains("{N}") == true)
                                            {
                                                col_bend_pos = colpt14;

                                            }
                                            if (check17.Contains("{O}") == true)
                                            {
                                                col_bend_pos = colpt15;

                                            }
                                            if (check17.Contains("{P}") == true)
                                            {
                                                col_bend_pos = colpt16;
                                            }
                                            if (check17.Contains("{Q}") == true)
                                            {
                                                col_bend_pos = colpt17;
                                            }
                                            if (check17.Contains("{R}") == true)
                                            {
                                                col_bend_pos = colpt18;
                                            }
                                            if (check17.Contains("{S}") == true)
                                            {
                                                col_bend_pos = colpt19;
                                            }
                                            if (check17.Contains("{T}") == true)
                                            {
                                                col_bend_pos = colpt20;
                                            }
                                            if (check17.Contains("{U}") == true)
                                            {
                                                col_bend_pos = colpt21;
                                            }
                                            if (check17.Contains("{V}") == true)
                                            {
                                                col_bend_pos = colpt22;
                                            }
                                            if (check17.Contains("{W}") == true)
                                            {
                                                col_bend_pos = colpt23;
                                            }
                                            if (check17.Contains("{X}") == true)
                                            {
                                                col_bend_pos = colpt24;
                                            }
                                            if (check17.Contains("{Y}") == true)
                                            {
                                                col_bend_pos = colpt25;
                                            }
                                            if (check17.Contains("{Z}") == true)
                                            {
                                                col_bend_pos = colpt26;
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
                                                col_bend_hor = colpt6;

                                            }
                                            if (check18.Contains("{G}") == true)
                                            {
                                                col_bend_hor = colpt7;

                                            }
                                            if (check18.Contains("{H}") == true)
                                            {
                                                col_bend_hor = colpt8;

                                            }
                                            if (check18.Contains("{I}") == true)
                                            {
                                                col_bend_hor = colpt9;

                                            }
                                            if (check18.Contains("{J}") == true)
                                            {
                                                col_bend_hor = colpt10;

                                            }
                                            if (check18.Contains("{K}") == true)
                                            {
                                                col_bend_hor = colpt11;

                                            }
                                            if (check18.Contains("{L}") == true)
                                            {
                                                col_bend_hor = colpt12;

                                            }
                                            if (check18.Contains("{M}") == true)
                                            {
                                                col_bend_hor = colpt13;

                                            }
                                            if (check18.Contains("{N}") == true)
                                            {
                                                col_bend_hor = colpt14;

                                            }
                                            if (check18.Contains("{O}") == true)
                                            {
                                                col_bend_hor = colpt15;

                                            }
                                            if (check18.Contains("{P}") == true)
                                            {
                                                col_bend_hor = colpt16;
                                            }

                                            if (check18.Contains("{Q}") == true)
                                            {
                                                col_bend_hor = colpt17;
                                            }
                                            if (check18.Contains("{R}") == true)
                                            {
                                                col_bend_hor = colpt18;
                                            }
                                            if (check18.Contains("{S}") == true)
                                            {
                                                col_bend_hor = colpt19;
                                            }
                                            if (check18.Contains("{T}") == true)
                                            {
                                                col_bend_hor = colpt20;
                                            }
                                            if (check18.Contains("{U}") == true)
                                            {
                                                col_bend_hor = colpt21;
                                            }
                                            if (check18.Contains("{V}") == true)
                                            {
                                                col_bend_hor = colpt22;
                                            }
                                            if (check18.Contains("{W}") == true)
                                            {
                                                col_bend_hor = colpt23;
                                            }
                                            if (check18.Contains("{X}") == true)
                                            {
                                                col_bend_hor = colpt24;
                                            }
                                            if (check18.Contains("{Y}") == true)
                                            {
                                                col_bend_hor = colpt25;
                                            }
                                            if (check18.Contains("{Z}") == true)
                                            {
                                                col_bend_hor = colpt26;
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
                                                col_bend_ver = colpt6;

                                            }
                                            if (check19.Contains("{G}") == true)
                                            {
                                                col_bend_ver = colpt7;

                                            }
                                            if (check19.Contains("{H}") == true)
                                            {
                                                col_bend_ver = colpt8;

                                            }
                                            if (check19.Contains("{I}") == true)
                                            {
                                                col_bend_ver = colpt9;

                                            }
                                            if (check19.Contains("{J}") == true)
                                            {
                                                col_bend_ver = colpt10;

                                            }
                                            if (check19.Contains("{K}") == true)
                                            {
                                                col_bend_ver = colpt11;

                                            }
                                            if (check19.Contains("{L}") == true)
                                            {
                                                col_bend_ver = colpt12;

                                            }
                                            if (check19.Contains("{M}") == true)
                                            {
                                                col_bend_ver = colpt13;

                                            }
                                            if (check19.Contains("{N}") == true)
                                            {
                                                col_bend_ver = colpt14;

                                            }
                                            if (check19.Contains("{O}") == true)
                                            {
                                                col_bend_ver = colpt15;

                                            }
                                            if (check19.Contains("{P}") == true)
                                            {
                                                col_bend_ver = colpt16;
                                            }

                                            if (check19.Contains("{Q}") == true)
                                            {
                                                col_bend_ver = colpt17;
                                            }
                                            if (check19.Contains("{R}") == true)
                                            {
                                                col_bend_ver = colpt18;
                                            }
                                            if (check19.Contains("{S}") == true)
                                            {
                                                col_bend_ver = colpt19;
                                            }
                                            if (check19.Contains("{T}") == true)
                                            {
                                                col_bend_ver = colpt20;
                                            }
                                            if (check19.Contains("{U}") == true)
                                            {
                                                col_bend_ver = colpt21;
                                            }
                                            if (check19.Contains("{V}") == true)
                                            {
                                                col_bend_ver = colpt22;
                                            }
                                            if (check19.Contains("{W}") == true)
                                            {
                                                col_bend_ver = colpt23;
                                            }
                                            if (check19.Contains("{X}") == true)
                                            {
                                                col_bend_ver = colpt24;
                                            }
                                            if (check19.Contains("{Y}") == true)
                                            {
                                                col_bend_ver = colpt25;
                                            }
                                            if (check19.Contains("{Z}") == true)
                                            {
                                                col_bend_ver = colpt26;
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
                                                col_elbow_type = colpt6;

                                            }
                                            if (check15.Contains("{G}") == true)
                                            {
                                                col_elbow_type = colpt7;

                                            }
                                            if (check15.Contains("{H}") == true)
                                            {
                                                col_elbow_type = colpt8;

                                            }
                                            if (check15.Contains("{I}") == true)
                                            {
                                                col_elbow_type = colpt9;

                                            }
                                            if (check15.Contains("{J}") == true)
                                            {
                                                col_elbow_type = colpt10;

                                            }
                                            if (check15.Contains("{K}") == true)
                                            {
                                                col_elbow_type = colpt11;

                                            }
                                            if (check15.Contains("{L}") == true)
                                            {
                                                col_elbow_type = colpt12;

                                            }
                                            if (check15.Contains("{M}") == true)
                                            {
                                                col_elbow_type = colpt13;
                                            }
                                            if (check15.Contains("{N}") == true)
                                            {
                                                col_elbow_type = colpt14;
                                            }
                                            if (check15.Contains("{O}") == true)
                                            {
                                                col_elbow_type = colpt15;
                                            }
                                            if (check15.Contains("{P}") == true)
                                            {
                                                col_elbow_type = colpt16;
                                            }

                                            if (check15.Contains("{Q}") == true)
                                            {
                                                col_elbow_type = colpt17;
                                            }
                                            if (check15.Contains("{R}") == true)
                                            {
                                                col_elbow_type = colpt18;
                                            }
                                            if (check15.Contains("{S}") == true)
                                            {
                                                col_elbow_type = colpt19;
                                            }
                                            if (check15.Contains("{T}") == true)
                                            {
                                                col_elbow_type = colpt20;
                                            }
                                            if (check15.Contains("{U}") == true)
                                            {
                                                col_elbow_type = colpt21;
                                            }
                                            if (check15.Contains("{V}") == true)
                                            {
                                                col_elbow_type = colpt22;
                                            }
                                            if (check15.Contains("{W}") == true)
                                            {
                                                col_elbow_type = colpt23;
                                            }
                                            if (check15.Contains("{X}") == true)
                                            {
                                                col_elbow_type = colpt24;
                                            }
                                            if (check15.Contains("{Y}") == true)
                                            {
                                                col_elbow_type = colpt25;
                                            }
                                            if (check15.Contains("{Z}") == true)
                                            {
                                                col_elbow_type = colpt26;
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
                                                col_elbow_defl = colpt6;

                                            }
                                            if (check16.Contains("{G}") == true)
                                            {
                                                col_elbow_defl = colpt7;

                                            }
                                            if (check16.Contains("{H}") == true)
                                            {
                                                col_elbow_defl = colpt8;

                                            }
                                            if (check16.Contains("{I}") == true)
                                            {
                                                col_elbow_defl = colpt9;

                                            }
                                            if (check16.Contains("{J}") == true)
                                            {
                                                col_elbow_defl = colpt10;

                                            }
                                            if (check16.Contains("{K}") == true)
                                            {
                                                col_elbow_defl = colpt11;

                                            }
                                            if (check16.Contains("{L}") == true)
                                            {
                                                col_elbow_defl = colpt12;

                                            }
                                            if (check16.Contains("{M}") == true)
                                            {
                                                col_elbow_defl = colpt13;

                                            }
                                            if (check16.Contains("{N}") == true)
                                            {
                                                col_elbow_defl = colpt14;

                                            }
                                            if (check16.Contains("{O}") == true)
                                            {
                                                col_elbow_defl = colpt15;

                                            }
                                            if (check16.Contains("{P}") == true)
                                            {
                                                col_elbow_defl = colpt16;

                                            }
                                            if (check16.Contains("{Q}") == true)
                                            {
                                                col_elbow_defl = colpt17;
                                            }
                                            if (check16.Contains("{R}") == true)
                                            {
                                                col_elbow_defl = colpt18;
                                            }
                                            if (check16.Contains("{S}") == true)
                                            {
                                                col_elbow_defl = colpt19;
                                            }
                                            if (check16.Contains("{T}") == true)
                                            {
                                                col_elbow_defl = colpt20;
                                            }
                                            if (check16.Contains("{U}") == true)
                                            {
                                                col_elbow_defl = colpt21;
                                            }
                                            if (check16.Contains("{V}") == true)
                                            {
                                                col_elbow_defl = colpt22;
                                            }
                                            if (check16.Contains("{W}") == true)
                                            {
                                                col_elbow_defl = colpt23;
                                            }
                                            if (check16.Contains("{X}") == true)
                                            {
                                                col_elbow_defl = colpt24;
                                            }
                                            if (check16.Contains("{Y}") == true)
                                            {
                                                col_elbow_defl = colpt25;
                                            }
                                            if (check16.Contains("{Z}") == true)
                                            {
                                                col_elbow_defl = colpt26;
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
                                                col_elbow_pos = colpt6;

                                            }
                                            if (check17.Contains("{G}") == true)
                                            {
                                                col_elbow_pos = colpt7;

                                            }
                                            if (check17.Contains("{H}") == true)
                                            {
                                                col_elbow_pos = colpt8;

                                            }
                                            if (check17.Contains("{I}") == true)
                                            {
                                                col_elbow_pos = colpt9;

                                            }
                                            if (check17.Contains("{J}") == true)
                                            {
                                                col_elbow_pos = colpt10;

                                            }
                                            if (check17.Contains("{K}") == true)
                                            {
                                                col_elbow_pos = colpt11;

                                            }
                                            if (check17.Contains("{L}") == true)
                                            {
                                                col_elbow_pos = colpt12;

                                            }
                                            if (check17.Contains("{M}") == true)
                                            {
                                                col_elbow_pos = colpt13;

                                            }
                                            if (check17.Contains("{N}") == true)
                                            {
                                                col_elbow_pos = colpt14;

                                            }
                                            if (check17.Contains("{O}") == true)
                                            {
                                                col_elbow_pos = colpt15;

                                            }
                                            if (check17.Contains("{P}") == true)
                                            {
                                                col_elbow_pos = colpt16;
                                            }
                                            if (check17.Contains("{Q}") == true)
                                            {
                                                col_elbow_pos = colpt17;
                                            }
                                            if (check17.Contains("{R}") == true)
                                            {
                                                col_elbow_pos = colpt18;
                                            }
                                            if (check17.Contains("{S}") == true)
                                            {
                                                col_elbow_pos = colpt19;
                                            }
                                            if (check17.Contains("{T}") == true)
                                            {
                                                col_elbow_pos = colpt20;
                                            }
                                            if (check17.Contains("{U}") == true)
                                            {
                                                col_elbow_pos = colpt21;
                                            }
                                            if (check17.Contains("{V}") == true)
                                            {
                                                col_elbow_pos = colpt22;
                                            }
                                            if (check17.Contains("{W}") == true)
                                            {
                                                col_elbow_pos = colpt23;
                                            }
                                            if (check17.Contains("{X}") == true)
                                            {
                                                col_elbow_pos = colpt24;
                                            }
                                            if (check17.Contains("{Y}") == true)
                                            {
                                                col_elbow_pos = colpt25;
                                            }
                                            if (check17.Contains("{Z}") == true)
                                            {
                                                col_elbow_pos = colpt26;
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
                                                col_elbow_hor = colpt6;

                                            }
                                            if (check18.Contains("{G}") == true)
                                            {
                                                col_elbow_hor = colpt7;

                                            }
                                            if (check18.Contains("{H}") == true)
                                            {
                                                col_elbow_hor = colpt8;

                                            }
                                            if (check18.Contains("{I}") == true)
                                            {
                                                col_elbow_hor = colpt9;

                                            }
                                            if (check18.Contains("{J}") == true)
                                            {
                                                col_elbow_hor = colpt10;

                                            }
                                            if (check18.Contains("{K}") == true)
                                            {
                                                col_elbow_hor = colpt11;

                                            }
                                            if (check18.Contains("{L}") == true)
                                            {
                                                col_elbow_hor = colpt12;

                                            }
                                            if (check18.Contains("{M}") == true)
                                            {
                                                col_elbow_hor = colpt13;

                                            }
                                            if (check18.Contains("{N}") == true)
                                            {
                                                col_elbow_hor = colpt14;

                                            }
                                            if (check18.Contains("{O}") == true)
                                            {
                                                col_elbow_hor = colpt15;

                                            }
                                            if (check18.Contains("{P}") == true)
                                            {
                                                col_elbow_hor = colpt16;
                                            }

                                            if (check18.Contains("{Q}") == true)
                                            {
                                                col_elbow_hor = colpt17;
                                            }
                                            if (check18.Contains("{R}") == true)
                                            {
                                                col_elbow_hor = colpt18;
                                            }
                                            if (check18.Contains("{S}") == true)
                                            {
                                                col_elbow_hor = colpt19;
                                            }
                                            if (check18.Contains("{T}") == true)
                                            {
                                                col_elbow_hor = colpt20;
                                            }
                                            if (check18.Contains("{U}") == true)
                                            {
                                                col_elbow_hor = colpt21;
                                            }
                                            if (check18.Contains("{V}") == true)
                                            {
                                                col_elbow_hor = colpt22;
                                            }
                                            if (check18.Contains("{W}") == true)
                                            {
                                                col_elbow_hor = colpt23;
                                            }
                                            if (check18.Contains("{X}") == true)
                                            {
                                                col_elbow_hor = colpt24;
                                            }
                                            if (check18.Contains("{Y}") == true)
                                            {
                                                col_elbow_hor = colpt25;
                                            }
                                            if (check18.Contains("{Z}") == true)
                                            {
                                                col_elbow_hor = colpt26;
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
                                                col_elbow_ver = colpt6;

                                            }
                                            if (check19.Contains("{G}") == true)
                                            {
                                                col_elbow_ver = colpt7;

                                            }
                                            if (check19.Contains("{H}") == true)
                                            {
                                                col_elbow_ver = colpt8;

                                            }
                                            if (check19.Contains("{I}") == true)
                                            {
                                                col_elbow_ver = colpt9;

                                            }
                                            if (check19.Contains("{J}") == true)
                                            {
                                                col_elbow_ver = colpt10;

                                            }
                                            if (check19.Contains("{K}") == true)
                                            {
                                                col_elbow_ver = colpt11;

                                            }
                                            if (check19.Contains("{L}") == true)
                                            {
                                                col_elbow_ver = colpt12;

                                            }
                                            if (check19.Contains("{M}") == true)
                                            {
                                                col_elbow_ver = colpt13;

                                            }
                                            if (check19.Contains("{N}") == true)
                                            {
                                                col_elbow_ver = colpt14;

                                            }
                                            if (check19.Contains("{O}") == true)
                                            {
                                                col_elbow_ver = colpt15;

                                            }
                                            if (check19.Contains("{P}") == true)
                                            {
                                                col_elbow_ver = colpt16;
                                            }

                                            if (check19.Contains("{Q}") == true)
                                            {
                                                col_elbow_ver = colpt17;
                                            }
                                            if (check19.Contains("{R}") == true)
                                            {
                                                col_elbow_ver = colpt18;
                                            }
                                            if (check19.Contains("{S}") == true)
                                            {
                                                col_elbow_ver = colpt19;
                                            }
                                            if (check19.Contains("{T}") == true)
                                            {
                                                col_elbow_ver = colpt20;
                                            }
                                            if (check19.Contains("{U}") == true)
                                            {
                                                col_elbow_ver = colpt21;
                                            }
                                            if (check19.Contains("{V}") == true)
                                            {
                                                col_elbow_ver = colpt22;
                                            }
                                            if (check19.Contains("{W}") == true)
                                            {
                                                col_elbow_ver = colpt23;
                                            }
                                            if (check19.Contains("{X}") == true)
                                            {
                                                col_elbow_ver = colpt24;
                                            }
                                            if (check19.Contains("{Y}") == true)
                                            {
                                                col_elbow_ver = colpt25;
                                            }
                                            if (check19.Contains("{Z}") == true)
                                            {
                                                col_elbow_ver = colpt26;
                                            }
                                        }
                                        #endregion
                                    }

                                    #endregion
                                }

                                if (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "WELD" || Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "WLD")
                                {
                                    #region WELD
                                    if (Wgen_main_form.dt_feature_codes.Rows[j][20] != DBNull.Value)
                                    {
                                        #region mm back
                                        string check20 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][20]);
                                        if (check20.Contains("{F}") == true ||
                                            check20.Contains("{G}") == true ||
                                            check20.Contains("{H}") == true ||
                                            check20.Contains("{I}") == true ||
                                            check20.Contains("{J}") == true ||
                                            check20.Contains("{K}") == true ||
                                            check20.Contains("{L}") == true ||
                                            check20.Contains("{M}") == true ||
                                            check20.Contains("{N}") == true ||
                                            check20.Contains("{O}") == true ||
                                            check20.Contains("{P}") == true
                                            )
                                        {
                                            if (check20.Contains("{F}") == true)
                                            {
                                                col_mm_back = colpt6;
                                            }
                                            if (check20.Contains("{G}") == true)
                                            {
                                                col_mm_back = colpt7;
                                            }
                                            if (check20.Contains("{H}") == true)
                                            {
                                                col_mm_back = colpt8;
                                            }
                                            if (check20.Contains("{I}") == true)
                                            {
                                                col_mm_back = colpt9;
                                            }
                                            if (check20.Contains("{J}") == true)
                                            {
                                                col_mm_back = colpt10;
                                            }
                                            if (check20.Contains("{K}") == true)
                                            {
                                                col_mm_back = colpt11;
                                            }
                                            if (check20.Contains("{L}") == true)
                                            {
                                                col_mm_back = colpt12;
                                            }
                                            if (check20.Contains("{M}") == true)
                                            {
                                                col_mm_back = colpt13;
                                            }
                                            if (check20.Contains("{N}") == true)
                                            {
                                                col_mm_back = colpt14;
                                            }
                                            if (check20.Contains("{O}") == true)
                                            {
                                                col_mm_back = colpt15;
                                            }
                                            if (check20.Contains("{P}") == true)
                                            {
                                                col_mm_back = colpt16;

                                            }
                                        }
                                        #endregion

                                        #region mm ahead
                                        string check21 = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][21]);
                                        if (check21.Contains("{F}") == true ||
                                            check21.Contains("{G}") == true ||
                                            check21.Contains("{H}") == true ||
                                            check21.Contains("{I}") == true ||
                                            check21.Contains("{J}") == true ||
                                            check21.Contains("{K}") == true ||
                                            check21.Contains("{L}") == true ||
                                            check21.Contains("{M}") == true ||
                                            check21.Contains("{N}") == true ||
                                            check21.Contains("{O}") == true ||
                                            check21.Contains("{P}") == true
                                            )
                                        {
                                            if (check21.Contains("{F}") == true)
                                            {
                                                col_mm_ahead = colpt6;

                                            }
                                            if (check21.Contains("{G}") == true)
                                            {
                                                col_mm_ahead = colpt7;

                                            }
                                            if (check21.Contains("{H}") == true)
                                            {
                                                col_mm_ahead = colpt8;

                                            }
                                            if (check21.Contains("{I}") == true)
                                            {
                                                col_mm_ahead = colpt9;

                                            }
                                            if (check21.Contains("{J}") == true)
                                            {
                                                col_mm_ahead = colpt10;

                                            }
                                            if (check21.Contains("{K}") == true)
                                            {
                                                col_mm_ahead = colpt11;

                                            }
                                            if (check21.Contains("{L}") == true)
                                            {
                                                col_mm_ahead = colpt12;

                                            }
                                            if (check21.Contains("{M}") == true)
                                            {
                                                col_mm_ahead = colpt13;

                                            }
                                            if (check21.Contains("{N}") == true)
                                            {
                                                col_mm_ahead = colpt14;

                                            }
                                            if (check21.Contains("{O}") == true)
                                            {
                                                col_mm_ahead = colpt15;

                                            }
                                            if (check21.Contains("{P}") == true)
                                            {
                                                col_mm_ahead = colpt16;

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

                DataSet dataset1 = new DataSet();
                dataset1.Tables.Add(Wgen_main_form.dt_all_points);
                dataset1.Tables.Add(Wgen_main_form.dt_ground_tally);

                DataRelation relation_mmid_back = new DataRelation("xxx", Wgen_main_form.dt_all_points.Columns[col_mm_back], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
                dataset1.Relations.Add(relation_mmid_back);

                DataRelation relation_mmid_ahead = new DataRelation("xxx1", Wgen_main_form.dt_all_points.Columns[col_mm_ahead], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
                dataset1.Relations.Add(relation_mmid_ahead);



                //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Wgen_main_form.dt_all_points);

                for (i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                {
                    if (Wgen_main_form.dt_all_points.Rows[i][colpt1] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt5] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt_sta] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt_sta]).Replace("+", "")) == true &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt2] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2])) == true &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt3] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3])) == true &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt4] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4])) == true)
                    {
                        string Feature_code = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt5]);
                        string pt1 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt1]);

                        if (Wgen_main_form.lista_feature_code_exception == null || Wgen_main_form.lista_feature_code_exception.Count == 0 || Wgen_main_form.lista_feature_code_exception.Contains(Feature_code) == false)
                        {
                            if (Feature_code.ToUpper() == "NATURAL_GROUND")
                            {
                                dt_ng.Rows.Add();
                                dt_ng.Rows[dt_ng.Rows.Count - 1][0] = Wgen_main_form.dt_all_points.Rows[i][colpt1];
                                dt_ng.Rows[dt_ng.Rows.Count - 1][1] = new Point3d(Convert.ToDouble(Wgen_main_form.dt_all_points.Rows[i][colpt3]),
                                                                                   Convert.ToDouble(Wgen_main_form.dt_all_points.Rows[i][colpt2]),
                                                                                    Convert.ToDouble(Wgen_main_form.dt_all_points.Rows[i][colpt4]));
                            }
                            else
                            {
                                if (Feature_code.ToUpper() == "WELD" || Feature_code.ToUpper() == "WLD")
                                {

                                    Wgen_main_form.dt_weld_map.Rows.Add();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_feat_code] = Feature_code.ToUpper();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_sta_lin] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_sta_lin]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_sta_ifc] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_sta_ifc]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_pnt] = pt1;
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_y] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_x] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_z] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4]);

                                    #region FEATURE CODES MAPPING
                                    for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                                    {
                                        string fc = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]).ToUpper();
                                        if (Wgen_main_form.client_name == Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]) &&
                                            Feature_code.ToUpper() == fc && (bool)Wgen_main_form.dt_feature_codes.Rows[j][2] == true)
                                        {
                                            string descr = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][3]).ToUpper();

                                            string G = "";
                                            string H = "";
                                            string I = "";
                                            string J = "";
                                            string K = "";
                                            string L = "";
                                            string M = "";
                                            string N = "";
                                            string O = "";
                                            string P = "";

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt7] != DBNull.Value)
                                            {
                                                G = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt7]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt8] != DBNull.Value)
                                            {
                                                H = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt8]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt9] != DBNull.Value)
                                            {
                                                I = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt9]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt10] != DBNull.Value)
                                            {
                                                J = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt10]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt11] != DBNull.Value)
                                            {
                                                K = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt11]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt12] != DBNull.Value)
                                            {
                                                L = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt12]).ToUpper();
                                            }
                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt13] != DBNull.Value)
                                            {
                                                M = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt13]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt14] != DBNull.Value)
                                            {
                                                N = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt14]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt15] != DBNull.Value)
                                            {
                                                O = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt15]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt16] != DBNull.Value)
                                            {
                                                P = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt16]).ToUpper();
                                            }

                                            descr = descr.Replace("{G}", G).Replace("{H}", H).Replace("{I}", I).Replace("{J}", J).Replace("{K}", K).Replace("{L}", L).Replace("{M}", M).Replace("{N}", N).Replace("{O}", O).Replace("{P}", P);
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] = descr;
                                            j = Wgen_main_form.dt_feature_codes.Rows.Count;
                                        }
                                    }
                                    #endregion


                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt7] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_location] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt7]);
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt6] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_file_name] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt6]);

                                    dt_colors.Rows.Add();

                                    int nr_match = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation_mmid_back).Length;
                                    if (nr_match == 1)
                                    {
                                        System.Data.DataRow row1 = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation_mmid_back)[0];

                                        if (row1[colgt1] != DBNull.Value)
                                        {
                                            string mmid = Convert.ToString(row1[colgt1]);
                                            string pipeID = "";
                                            string heat = "";
                                            string wt = "";
                                            string diam = "";
                                            string grd = "";
                                            string coating = "";
                                            string dj_back = "";
                                            string man_back = "";
                                            string len_back = "";

                                            if (row1[colgt2] != DBNull.Value)
                                            {
                                                pipeID = Convert.ToString(row1[colgt2]);
                                            }
                                            if (row1[colgt3] != DBNull.Value)
                                            {
                                                heat = Convert.ToString(row1[colgt3]);
                                            }
                                            if (row1[colgt6] != DBNull.Value)
                                            {
                                                wt = Convert.ToString(row1[colgt6]);
                                            }
                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                diam = Convert.ToString(row1[colgt7]);
                                            }
                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                grd = Convert.ToString(row1[colgt8]);
                                            }

                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                coating = Convert.ToString(row1[colgt9]);
                                            }

                                            if (row1[colgt11] != DBNull.Value)
                                            {
                                                dj_back = Convert.ToString(row1[colgt11]);
                                            }

                                            if (checkBox_add_djnumber_to_pipeID.Checked == true)
                                            {
                                                if (dj_back != "")
                                                {
                                                    pipeID = pipeID + "/" + dj_back;
                                                }
                                            }

                                            if (row1[colgt10] != DBNull.Value)
                                            {
                                                man_back = Convert.ToString(row1[colgt10]);
                                            }


                                            if (row1[colgt4] != DBNull.Value)
                                            {
                                                len_back = Convert.ToString(row1[colgt4]);
                                            }
                                            if (row1[colgt5] != DBNull.Value)
                                            {
                                                len_back = Convert.ToString(row1[colgt5]);
                                            }

                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_mm_bk] = mmid;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_wall_bk] = wt;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_pipe_bk] = pipeID;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_heat_bk] = heat;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_coat_bk] = coating;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_grade_bk] = grd;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_manufacture_bk] = man_back;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_length_bk] = len_back;


                                        }
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation_mmid_ahead).Length == 1)
                                    {
                                        System.Data.DataRow row1 = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation_mmid_ahead)[0];
                                        if (row1[colgt1] != DBNull.Value)
                                        {
                                            string mmid = Convert.ToString(row1[colgt1]);
                                            string pipeID = "";
                                            string heat = "";
                                            string wt = "";
                                            string diam = "";
                                            string grd = "";
                                            string coating = "";
                                            string dj_ahead = "";
                                            string man_ahead = "";
                                            string len_ahead = "";

                                            if (row1[colgt2] != DBNull.Value)
                                            {
                                                pipeID = Convert.ToString(row1[colgt2]);
                                            }
                                            if (row1[colgt3] != DBNull.Value)
                                            {
                                                heat = Convert.ToString(row1[colgt3]);
                                            }
                                            if (row1[colgt6] != DBNull.Value)
                                            {
                                                wt = Convert.ToString(row1[colgt6]);
                                            }
                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                diam = Convert.ToString(row1[colgt7]);
                                            }
                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                grd = Convert.ToString(row1[colgt8]);
                                            }

                                            if (row1[colgt7] != DBNull.Value)
                                            {
                                                coating = Convert.ToString(row1[colgt9]);
                                            }
                                            if (row1[colgt11] != DBNull.Value)
                                            {
                                                dj_ahead = Convert.ToString(row1[colgt11]);
                                            }

                                            if (checkBox_add_djnumber_to_pipeID.Checked == true)
                                            {
                                                if (dj_ahead != "")
                                                {
                                                    pipeID = pipeID + "/" + dj_ahead;
                                                }
                                            }


                                            if (row1[colgt10] != DBNull.Value)
                                            {
                                                man_ahead = Convert.ToString(row1[colgt10]);
                                            }

                                            if (row1[colgt4] != DBNull.Value)
                                            {
                                                len_ahead = Convert.ToString(row1[colgt4]);
                                            }

                                            if (row1[colgt5] != DBNull.Value)
                                            {
                                                len_ahead = Convert.ToString(row1[colgt5]);
                                            }

                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_mm_ahd] = mmid;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_wall_ahd] = wt;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_pipe_ahd] = pipeID;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_heat_ahd] = heat;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_coat_ahd] = coating;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_grade_ahd] = grd;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_manufacture_ahd] = man_ahead;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_length_ahd] = len_ahead;

                                        }
                                    }
                                }

                                else if (Wgen_main_form.lista_feature_code.Contains(Feature_code.ToUpper()) == true)
                                {
                                    Wgen_main_form.dt_weld_map.Rows.Add();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_feat_code] = Feature_code.ToUpper();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_sta_lin] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_sta_lin]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_sta_ifc] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_sta_ifc]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_pnt] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt1]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_y] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_x] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_z] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4]);

                                    string descr = "NOT DEFINED";

                                    #region FEATURE CODES MAPPING
                                    for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                                    {
                                        if (Wgen_main_form.client_name == Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]) &&
                                            Feature_code.ToUpper() == Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]).ToUpper() &&
                                           (bool)Wgen_main_form.dt_feature_codes.Rows[j][2] == true)
                                        {
                                            descr = Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][3]).ToUpper();
                                            string G = "";
                                            string H = "";
                                            string I = "";
                                            string J = "";
                                            string K = "";
                                            string L = "";
                                            string M = "";
                                            string N = "";
                                            string O = "";
                                            string P = "";



                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt7] != DBNull.Value)
                                            {
                                                G = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt7]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt8] != DBNull.Value)
                                            {
                                                H = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt8]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt9] != DBNull.Value)
                                            {
                                                I = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt9]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt10] != DBNull.Value)
                                            {
                                                J = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt10]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt11] != DBNull.Value)
                                            {
                                                K = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt11]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt12] != DBNull.Value)
                                            {
                                                L = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt12]).ToUpper();
                                            }
                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt13] != DBNull.Value)
                                            {
                                                M = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt13]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt14] != DBNull.Value)
                                            {
                                                N = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt14]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt15] != DBNull.Value)
                                            {
                                                O = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt15]).ToUpper();
                                            }

                                            if (Wgen_main_form.dt_all_points.Rows[i][colpt16] != DBNull.Value)
                                            {
                                                P = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt16]).ToUpper();
                                            }


                                            descr = descr.Replace("{G}", G).Replace("{H}", H).Replace("{I}", I).Replace("{J}", J).Replace("{K}", K).Replace("{L}", L).Replace("{M}", M).Replace("{N}", N).Replace("{O}", O).Replace("{P}", P);
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] = descr;
                                            j = Wgen_main_form.dt_feature_codes.Rows.Count;
                                        }
                                    }
                                    #endregion

                                    if (Feature_code.ToUpper() == "BEND")
                                    {
                                        if (Wgen_main_form.dt_all_points.Rows[i][col_bend_defl] != DBNull.Value && Wgen_main_form.dt_all_points.Rows[i][col_bend_pos] != DBNull.Value)
                                        {
                                            string bend_type_field_induction = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_type]).ToUpper();
                                            string bend_deflection_left_right_sag_overbnd = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_defl]).ToUpper();
                                            string position = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_pos] != DBNull.Value) position = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_pos]).ToUpper();
                                            string hdefl = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_hor] != DBNull.Value) hdefl = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_hor]).ToUpper();
                                            string vdefl = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_bend_ver] != DBNull.Value) vdefl = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_bend_ver]).ToUpper();


                                            if (bend_deflection_left_right_sag_overbnd == "RIGHT" || bend_deflection_left_right_sag_overbnd == "LEFT")
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                    bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR";
                                            }
                                            else if (bend_deflection_left_right_sag_overbnd == "SAG" || bend_deflection_left_right_sag_overbnd == "OVERBEND")
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                    bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + vdefl + " VER";
                                            }
                                            else
                                            {
                                                if (descr != "TOTAL ANGLE")
                                                {
                                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                        bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR" + "/" + vdefl + " VER";
                                                }
                                                else
                                                {
                                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                        bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " TOTAL";
                                                }

                                            }
                                        }
                                    }

                                    if (Feature_code.ToUpper() == "ELBOW")
                                    {
                                        #region ELBOW DESCRIPTION BUILT UP
                                        if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_defl] != DBNull.Value && Wgen_main_form.dt_all_points.Rows[i][col_elbow_pos] != DBNull.Value)
                                        {
                                            string ellbow_type_field_induction = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_type]).ToUpper();
                                            string ellbow_deflection_left_right_sag_overbnd = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_defl]).ToUpper();
                                            string position = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_pos] != DBNull.Value) position = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_pos]).ToUpper();
                                            string hdefl = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_hor] != DBNull.Value) hdefl = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_hor]).ToUpper();
                                            string vdefl = "NOT DEFINED";
                                            if (Wgen_main_form.dt_all_points.Rows[i][col_elbow_ver] != DBNull.Value) vdefl = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col_elbow_ver]).ToUpper();


                                            if (ellbow_deflection_left_right_sag_overbnd == "RIGHT" || ellbow_deflection_left_right_sag_overbnd == "LEFT")
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                    ellbow_type_field_induction + "/" + ellbow_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR";
                                            }
                                            else if (ellbow_deflection_left_right_sag_overbnd == "SAG" || ellbow_deflection_left_right_sag_overbnd == "OVERBEND")
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                    ellbow_type_field_induction + "/" + ellbow_deflection_left_right_sag_overbnd + "/" + position + "/" + vdefl + " VER";
                                            }
                                            else
                                            {
                                                if (descr != "TOTAL ANGLE")
                                                {
                                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                    ellbow_type_field_induction + "/" + ellbow_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR" + "/" + vdefl + " VER";
                                                }
                                                else
                                                {
                                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_descr] =
                                                        ellbow_type_field_induction + "/" + ellbow_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " TOTAL";
                                                }
                                            }
                                        }
                                        #endregion
                                    }


                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt6] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col_file_name] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt6]);

                                    string valpt8 = "";
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt8] != DBNull.Value) valpt8 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt8]);

                                    string valpt10 = "";
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt10] != DBNull.Value) valpt10 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt10]);

                                    string valpt11 = "";
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt11] != DBNull.Value) valpt11 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt11]);

                                    string valpt12 = "";
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt12] != DBNull.Value) valpt12 = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt12]);

                                    dt_colors.Rows.Add();


                                    if (Feature_code.ToUpper() == "LOOSE_END")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 39;
                                    }
                                    if (Feature_code.ToUpper() == "TRENCH_BREAKERS")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 15;
                                    }
                                    if (Feature_code.ToUpper() == "BEND" || Feature_code.ToUpper() == "ELBOW")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 37;
                                    }
                                    if (Feature_code.ToUpper() == "COATING_CHANGE")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 19;
                                    }
                                    if (Feature_code.ToUpper() == "CP_CAD_WELD" || Feature_code.ToUpper() == "CAD_WELD")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 7; //magenta
                                    }
                                    if (Feature_code.ToUpper() == "RIVER_WEIGHT" || Feature_code.ToUpper() == "WEIGHT")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 5;//blue
                                    }
                                    if (Feature_code.ToUpper() == "TRENCH_BREAKERS")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 44;//orange
                                    }
                                    if (Feature_code.ToUpper() == "ROCK_SHIELD")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 45;//tan
                                    }
                                    if (Feature_code.ToUpper() == "CENTERLINE_ROAD" || Feature_code.ToUpper() == "CENTERLINE_OF_ROAD")
                                    {
                                        dt_colors.Rows[dt_colors.Rows.Count - 1][1] = 4; //green
                                    }

                                }
                            }
                        }
                    }
                }


                dataset1.Relations.Remove(relation_mmid_back);
                dataset1.Relations.Remove(relation_mmid_ahead);
                dataset1.Tables.Remove(Wgen_main_form.dt_all_points);
                dataset1.Tables.Remove(Wgen_main_form.dt_ground_tally);

                if (dt_ng.Rows.Count > 0)
                {
                    //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_ng);

                    Wgen_main_form.dt_weld_map.Columns.Add("Ground Point to Weld Point Distance", typeof(double));
                    for (i = 0; i < Wgen_main_form.dt_weld_map.Rows.Count; ++i)
                    {

                        string pt_weld = Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col_pnt]);

                        Point3d pt1 = new Point3d(Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col_x]), Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col_y]), 0);

                        string ptg = "not available";
                        double Xmin = -1;
                        double Ymin = -1;
                        double Zmin = -1;
                        double d1 = ng_tolerance;

                        for (int j = 0; j < dt_ng.Rows.Count; ++j)
                        {
                            Point3d pt2 = (Point3d)dt_ng.Rows[j][1];



                            string pt_g = Convert.ToString(dt_ng.Rows[j][0]);
                            double d2 = Math.Pow(Math.Pow(pt1.X - pt2.X, 2) + Math.Pow(pt1.Y - pt2.Y, 2), 0.5);
                            if (d2 < d1)
                            {

                                d1 = d2;
                                ptg = pt_g;
                                Xmin = pt2.X;
                                Ymin = pt2.Y;
                                Zmin = pt2.Z;
                            }
                        }



                        Wgen_main_form.dt_weld_map.Rows[i]["Ground Point to Weld Point Distance"] = Math.Round(d1, 2);
                        Wgen_main_form.dt_weld_map.Rows[i][col_ng] = ptg;
                        Wgen_main_form.dt_weld_map.Rows[i][col_ng_y] = Convert.ToString(Ymin);
                        Wgen_main_form.dt_weld_map.Rows[i][col_ng_x] = Convert.ToString(Xmin);
                        Wgen_main_form.dt_weld_map.Rows[i][col_ng_z] = Convert.ToString(Zmin);
                        if (Wgen_main_form.dt_weld_map.Rows[i][col_z] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[i][col_cover] = Convert.ToString(Zmin - Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col_z]));



                    }
                }


                if (Wgen_main_form.dt_weld_map.Rows.Count > 0)
                {
                    wm_checks(Wgen_main_form.dt_weld_map);
                    W2 = Functions.Transfer_weldmap_datatable_to_new_excel_spreadsheet_formated_general_and_colored(Wgen_main_form.dt_weld_map, dt_colors);
                    button_refresh_ws1_Click(sender, e);
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(i.ToString() + ":\r\n" + ex.Message);
            }
            set_enable_true();

            Wgen_main_form.tpage_weldmap.Show();
            Wgen_main_form.tpage_blank.Hide();
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_blank.get_label_wait_visible(false);

        }

        private void transfer_weld_coordinates_to_pmc(System.Data.DataTable dtwm)
        {
            System.Data.DataTable dt_welds = new System.Data.DataTable();
            dt_welds.Columns.Add(col_pnt, typeof(string));
            dt_welds.Columns.Add(col_y, typeof(double));
            dt_welds.Columns.Add(col_x, typeof(double));
            dt_welds.Columns.Add(col_z, typeof(double));
            dt_welds.Columns.Add(col_sta_lin, typeof(string));
            dt_welds.Columns.Add(col_sta_ifc, typeof(string));

            System.Data.DataTable dt_pmc = new System.Data.DataTable();
            dt_pmc = dt_welds.Clone();
            dt_pmc.Columns.Add("index", typeof(int));

            for (int i = 0; i < dtwm.Rows.Count; ++i)
            {
                string feature1 = "xx";
                if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                {
                    feature1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                }

                if (feature1.ToUpper() == "PIPE_MATERIAL_CHANGE" || feature1.ToUpper() == "COATING_CHANGE" || feature1.ToUpper() == "PIP" || feature1.ToUpper() == "CC")
                {
                    dt_pmc.Rows.Add();
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1]["index"] = i;
                }

                if (feature1.ToUpper() == "WELD" || feature1.ToUpper() == "WLD")
                {
                    dt_welds.Rows.Add();
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                }
            }


            for (int i = 0; i < dt_pmc.Rows.Count; ++i)
            {
                double dmax = 0.5;
                if (dt_pmc.Rows[i][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[i][col_y])) == true &&
                    dt_pmc.Rows[i][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[i][col_x])) == true)
                {
                    double y1 = Convert.ToDouble(dt_pmc.Rows[i][col_y]);
                    double x1 = Convert.ToDouble(dt_pmc.Rows[i][col_x]);
                    for (int j = 0; j < dt_welds.Rows.Count; ++j)
                    {
                        if (dt_welds.Rows[j][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds.Rows[j][col_y])) == true &&
                            dt_welds.Rows[j][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds.Rows[j][col_x])) == true)
                        {
                            double y2 = Convert.ToDouble(dt_welds.Rows[j][col_y]);
                            double x2 = Convert.ToDouble(dt_welds.Rows[j][col_x]);
                            double d1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                            if (d1 < dmax)
                            {
                                dmax = d1;
                                dt_pmc.Rows[i][col_y] = dt_welds.Rows[j][col_y];
                                dt_pmc.Rows[i][col_x] = dt_welds.Rows[j][col_x];
                                dt_pmc.Rows[i][col_z] = dt_welds.Rows[j][col_z];
                                dt_pmc.Rows[i][col_sta_lin] = dt_welds.Rows[j][col_sta_lin];
                                dt_pmc.Rows[i][col_sta_ifc] = dt_welds.Rows[j][col_sta_ifc];
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < dt_pmc.Rows.Count; ++i)
            {
                int idx = Convert.ToInt32(dt_pmc.Rows[i]["index"]);
                dtwm.Rows[idx][col_y] = dt_pmc.Rows[i][col_y];
                dtwm.Rows[idx][col_x] = dt_pmc.Rows[i][col_x];
                dtwm.Rows[idx][col_z] = dt_pmc.Rows[i][col_z];
                dtwm.Rows[idx][col_sta_lin] = dt_pmc.Rows[i][col_sta_lin];
                dtwm.Rows[idx][col_sta_ifc] = dt_pmc.Rows[i][col_sta_ifc];

            }

        }

        private void build_dt_welds_pmc(System.Data.DataTable dtwm)
        {
            dt_welds_with_pmc = new System.Data.DataTable();
            dt_welds_with_pmc.Columns.Add(col_pnt, typeof(string));
            dt_welds_with_pmc.Columns.Add(col_y, typeof(double));
            dt_welds_with_pmc.Columns.Add(col_x, typeof(double));
            dt_welds_with_pmc.Columns.Add(col_z, typeof(double));
            dt_welds_with_pmc.Columns.Add(col_sta_lin, typeof(string));
            dt_welds_with_pmc.Columns.Add(col_sta_ifc, typeof(string));
            dt_welds_with_pmc.Columns.Add("PMC", typeof(string));
            dt_welds_with_pmc.Columns.Add("CC", typeof(string));

            System.Data.DataTable dt_pmc = new System.Data.DataTable();
            dt_pmc.Columns.Add(col_pnt, typeof(string));
            dt_pmc.Columns.Add(col_y, typeof(double));
            dt_pmc.Columns.Add(col_x, typeof(double));
            dt_pmc.Columns.Add(col_z, typeof(double));
            dt_pmc.Columns.Add(col_feat_code, typeof(string));
            dt_pmc.Columns.Add(col_sta_lin, typeof(string));
            dt_pmc.Columns.Add(col_sta_ifc, typeof(string));


            System.Data.DataTable dt_cc = new System.Data.DataTable();
            dt_cc.Columns.Add(col_pnt, typeof(string));
            dt_cc.Columns.Add(col_y, typeof(double));
            dt_cc.Columns.Add(col_x, typeof(double));
            dt_cc.Columns.Add(col_z, typeof(double));
            dt_cc.Columns.Add(col_feat_code, typeof(string));
            dt_cc.Columns.Add(col_sta_lin, typeof(string));
            dt_cc.Columns.Add(col_sta_ifc, typeof(string));

            System.Data.DataTable dt_le = new System.Data.DataTable();
            dt_le.Columns.Add(col_pnt, typeof(string));
            dt_le.Columns.Add(col_y, typeof(double));
            dt_le.Columns.Add(col_x, typeof(double));
            dt_le.Columns.Add(col_z, typeof(double));
            dt_le.Columns.Add(col_sta_lin, typeof(string));
            dt_le.Columns.Add(col_sta_ifc, typeof(string));

            for (int i = 0; i < dtwm.Rows.Count; ++i)
            {
                string feature1 = "xx";
                if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                {
                    feature1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                }

                if (feature1.ToUpper() == "PIPE_MATERIAL_CHANGE" || feature1.ToUpper() == "PIP")
                {
                    dt_pmc.Rows.Add();
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_feat_code] = dtwm.Rows[i][col_feat_code];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                }


                if (feature1.ToUpper() == "COATING_CHANGE" || feature1.ToUpper() == "CC")
                {
                    dt_cc.Rows.Add();
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_feat_code] = dtwm.Rows[i][col_feat_code];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_cc.Rows[dt_cc.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                }


                if (feature1.ToUpper() == "WELD" || feature1.ToUpper() == "WLD")
                {
                    dt_welds_with_pmc.Rows.Add();
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_welds_with_pmc.Rows[dt_welds_with_pmc.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                }

                if (feature1.ToUpper() == "LOOSE_END" || feature1.ToUpper() == "LE")
                {
                    dt_le.Rows.Add();
                    dt_le.Rows[dt_le.Rows.Count - 1][col_pnt] = dtwm.Rows[i][col_pnt];
                    dt_le.Rows[dt_le.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                    dt_le.Rows[dt_le.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                    dt_le.Rows[dt_le.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                    dt_le.Rows[dt_le.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                    dt_le.Rows[dt_le.Rows.Count - 1][col_sta_ifc] = dtwm.Rows[i][col_sta_ifc];
                }
            }

            for (int i = 0; i < dt_welds_with_pmc.Rows.Count; ++i)
            {
                if (dt_welds_with_pmc.Rows[i][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds_with_pmc.Rows[i][col_y])) == true &&
                    dt_welds_with_pmc.Rows[i][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds_with_pmc.Rows[i][col_x])) == true)
                {
                    double y1 = Convert.ToDouble(dt_welds_with_pmc.Rows[i][col_y]);
                    double x1 = Convert.ToDouble(dt_welds_with_pmc.Rows[i][col_x]);

                    for (int j = 0; j < dt_pmc.Rows.Count; ++j)
                    {
                        if (dt_pmc.Rows[j][col_pnt] != DBNull.Value &&
                            dt_pmc.Rows[j][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[j][col_y])) == true &&
                            dt_pmc.Rows[j][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[j][col_x])) == true)
                        {

                            string pt1 = Convert.ToString(dt_pmc.Rows[j][col_pnt]);
                            double y2 = Convert.ToDouble(dt_pmc.Rows[j][col_y]);
                            double x2 = Convert.ToDouble(dt_pmc.Rows[j][col_x]);
                            double d1 = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);
                            if (d1 < pmc_tolerance)
                            {
                                dt_welds_with_pmc.Rows[i]["PMC"] = pt1;
                                j = dt_pmc.Rows.Count;
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < dt_welds_with_pmc.Rows.Count; ++i)
            {
                if (dt_welds_with_pmc.Rows[i][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds_with_pmc.Rows[i][col_y])) == true &&
                    dt_welds_with_pmc.Rows[i][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds_with_pmc.Rows[i][col_x])) == true)
                {
                    double y1 = Convert.ToDouble(dt_welds_with_pmc.Rows[i][col_y]);
                    double x1 = Convert.ToDouble(dt_welds_with_pmc.Rows[i][col_x]);

                    for (int j = 0; j < dt_cc.Rows.Count; ++j)
                    {
                        if (dt_cc.Rows[j][col_pnt] != DBNull.Value &&
                            dt_cc.Rows[j][col_y] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_cc.Rows[j][col_y])) == true &&
                            dt_cc.Rows[j][col_x] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_cc.Rows[j][col_x])) == true)
                        {

                            string pt1 = Convert.ToString(dt_cc.Rows[j][col_pnt]);
                            double y2 = Convert.ToDouble(dt_cc.Rows[j][col_y]);
                            double x2 = Convert.ToDouble(dt_cc.Rows[j][col_x]);
                            double d1 = Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);
                            if (d1 < pmc_tolerance)
                            {
                                dt_welds_with_pmc.Rows[i]["CC"] = pt1;
                                j = dt_cc.Rows.Count;
                            }
                        }
                    }
                }
            }


            for (int i = 1; i < dt_welds_with_pmc.Rows.Count; ++i)
            {
                if (dt_welds_with_pmc.Rows[i - 1][col_sta_lin] != DBNull.Value && dt_welds_with_pmc.Rows[i][col_sta_lin] != DBNull.Value)
                {
                    string sta_string1 = Convert.ToString(dt_welds_with_pmc.Rows[i - 1][col_sta_lin]);
                    string sta_string2 = Convert.ToString(dt_welds_with_pmc.Rows[i][col_sta_lin]);
                    if (Functions.IsNumeric(Convert.ToString(sta_string1).Replace("+", "")) == true && Functions.IsNumeric(Convert.ToString(sta_string2).Replace("+", "")) == true)
                    {
                        double sta1 = Convert.ToDouble(sta_string1.Replace("+", ""));
                        double sta2 = Convert.ToDouble(sta_string2.Replace("+", ""));

                        for (int j = 0; j < dt_le.Rows.Count; ++j)
                        {
                            if (dt_le.Rows[j][col_sta_lin] != DBNull.Value)
                            {
                                string sta_string = Convert.ToString(dt_le.Rows[j][col_sta_lin]);
                                if (Functions.IsNumeric(Convert.ToString(sta_string).Replace("+", "")) == true)
                                {
                                    double sta = Convert.ToDouble(sta_string.Replace("+", ""));
                                    if (sta <= sta2 && sta >= sta1)
                                    {
                                        dt_welds_with_pmc.Rows[i]["PMC"] = "LE BACK";
                                        dt_welds_with_pmc.Rows[i - 1]["PMC"] = "LE AHEAD";
                                        dt_welds_with_pmc.Rows[i]["CC"] = "LE BACK";
                                        dt_welds_with_pmc.Rows[i - 1]["CC"] = "LE AHEAD";
                                        j = dt_le.Rows.Count;
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }


        private void wm_checks(System.Data.DataTable dtwm)
        {

            build_dt_welds_pmc(dtwm);

            int nr_duplicates = 0;
            int nr_null_values = 0;
            dt_errors = new System.Data.DataTable();
            dt_errors.Columns.Add("Point", typeof(string));
            dt_errors.Columns.Add("Feature Code", typeof(string));
            dt_errors.Columns.Add("Value", typeof(string));
            dt_errors.Columns.Add("Excel address", typeof(string));
            dt_errors.Columns.Add("Error type", typeof(string));
            dt_errors.Columns.Add("x", typeof(string));
            dt_errors.Columns.Add("y", typeof(string));


            var duplicates_pts = dtwm.AsEnumerable().GroupBy(datarow1 => new { pnt = datarow1.Field<string>(col_pnt) }).Where(g => g.Count() > 1).Select(g => new { g.Key.pnt }).ToList();
            var duplicates_mmid_back = dtwm.AsEnumerable().GroupBy(datarow1 => new { mmid = datarow1.Field<string>(col_mm_bk) }).Where(g => g.Count() > 1).Select(g => new { g.Key.mmid }).ToList();
            var duplicates_mmid_ahead = dtwm.AsEnumerable().GroupBy(datarow1 => new { mmid = datarow1.Field<string>(col_mm_ahd) }).Where(g => g.Count() > 1).Select(g => new { g.Key.mmid }).ToList();

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Columns.Add(col_pnt, typeof(string));
            dt2.TableName = "dt2";

            System.Data.DataTable dt_welds = new System.Data.DataTable();
            dt_welds.Columns.Add(col_pnt, typeof(string));
            dt_welds.Columns.Add(col_y, typeof(double));
            dt_welds.Columns.Add(col_x, typeof(double));
            dt_welds.Columns.Add(col_z, typeof(double));
            dt_welds.Columns.Add(col_sta_lin, typeof(string));

            dt_welds.Columns.Add(col_wall_bk, typeof(string)); //wall back
            dt_welds.Columns.Add(col_wall_ahd, typeof(string)); //wall ahead
            dt_welds.Columns.Add(col_coat_bk, typeof(string)); //coating back
            dt_welds.Columns.Add(col_coat_ahd, typeof(string)); //coating ahead
            dt_welds.Columns.Add("address", typeof(string));



            System.Data.DataTable dt_bend_welds = new System.Data.DataTable();
            dt_bend_welds.Columns.Add(col_pnt, typeof(string));
            dt_bend_welds.Columns.Add(col_y, typeof(double));
            dt_bend_welds.Columns.Add(col_x, typeof(double));
            dt_bend_welds.Columns.Add(col_z, typeof(double));
            dt_bend_welds.Columns.Add(col_mm_ahd, typeof(string));
            dt_bend_welds.Columns.Add("feature_code", typeof(string));
            dt_bend_welds.Columns.Add("length", typeof(double));
            dt_bend_welds.Columns.Add("address", typeof(string));


            DataSet dataset1 = new DataSet();
            dataset1.Tables.Add(dtwm);

            if (duplicates_pts.Count > 0)
            {
                for (int i = 0; i < duplicates_pts.Count; ++i)
                {
                    if (duplicates_pts[i].pnt != null)
                    {
                        string duplicat_val1 = Convert.ToString(duplicates_pts[i].pnt);
                        dt2.Rows.Add();
                        dt2.Rows[dt2.Rows.Count - 1][0] = duplicat_val1;
                    }
                }

                dataset1.Tables.Add(dt2);
                DataRelation relation1 = new DataRelation("xxx1", dtwm.Columns[col_pnt], dt2.Columns[col_pnt], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate points
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string Point1 = dtwm.Rows[i].GetChildRows(relation1)[j][col_pnt].ToString();

                            dt_errors.Rows.Add();
                            dt_errors.Rows[dt_errors.Rows.Count - 1][0] = Point1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][1] = Feature1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_1.Text + Convert.ToString(i + start_row);
                            dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Duplicate point number";
                            dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;


                        }
                    }
                    #endregion
                }

                dataset1.Relations.Remove(relation1);
                dataset1.Tables.Remove(dt2);
                dt2 = null;
            }



            #region mmid back duplicates
            dt2 = new System.Data.DataTable();
            dt2.Columns.Add(col_mm_bk, typeof(string));
            dt2.TableName = "dt2";

            if (duplicates_mmid_back.Count > 0)
            {
                for (int i = 0; i < duplicates_mmid_back.Count; ++i)
                {
                    if (duplicates_mmid_back[i].mmid != null)
                    {
                        string duplicat_val1 = Convert.ToString(duplicates_mmid_back[i].mmid);
                        dt2.Rows.Add();
                        dt2.Rows[dt2.Rows.Count - 1][0] = duplicat_val1;
                    }
                }

                dataset1.Tables.Add(dt2);
                DataRelation relation1 = new DataRelation("xxx9", dtwm.Columns[col_mm_bk], dt2.Columns[col_mm_bk], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = nr_duplicates + dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate mmid
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string pt1 = "xx";
                        if (dtwm.Rows[i][col_pnt] != DBNull.Value)
                        {
                            pt1 = Convert.ToString(dtwm.Rows[i][col_pnt]);
                        }
                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string mmid1 = dtwm.Rows[i].GetChildRows(relation1)[j][col_mm_bk].ToString();

                            dt_errors.Rows.Add();
                            dt_errors.Rows[dt_errors.Rows.Count - 1][0] = pt1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][1] = Feature1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][2] = mmid1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_9.Text + Convert.ToString(i + start_row);
                            dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Duplicate mmid back";
                            dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;


                        }
                    }
                    #endregion
                }

                dataset1.Relations.Remove(relation1);
                dataset1.Tables.Remove(dt2);
                dt2 = null;
            }
            #endregion

            #region mmid ahead duplicates
            dt2 = new System.Data.DataTable();
            dt2.Columns.Add(col_mm_ahd, typeof(string));
            dt2.TableName = "dt2";

            if (duplicates_mmid_ahead.Count > 0)
            {
                for (int i = 0; i < duplicates_mmid_ahead.Count; ++i)
                {
                    if (duplicates_mmid_ahead[i].mmid != null)
                    {
                        string duplicat_val1 = Convert.ToString(duplicates_mmid_ahead[i].mmid);
                        dt2.Rows.Add();
                        dt2.Rows[dt2.Rows.Count - 1][0] = duplicat_val1;
                    }
                }

                dataset1.Tables.Add(dt2);
                DataRelation relation1 = new DataRelation("xxx9", dtwm.Columns[col_mm_ahd], dt2.Columns[col_mm_ahd], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = nr_duplicates + dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate mmid
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string pt1 = "xx";
                        if (dtwm.Rows[i][col_pnt] != DBNull.Value)
                        {
                            pt1 = Convert.ToString(dtwm.Rows[i][col_pnt]);
                        }

                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string mmid1 = dtwm.Rows[i].GetChildRows(relation1)[j][col_mm_ahd].ToString();

                            dt_errors.Rows.Add();
                            dt_errors.Rows[dt_errors.Rows.Count - 1][0] = pt1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][1] = Feature1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][2] = mmid1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_17.Text + Convert.ToString(i + start_row);
                            dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "Duplicate mmid ahead";
                            dt_errors.Rows[dt_errors.Rows.Count - 1][5] = x;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][6] = y;


                        }
                    }
                    #endregion
                }

                dataset1.Relations.Remove(relation1);
                dataset1.Tables.Remove(dt2);
                dt2 = null;
            }
            #endregion


            List<string> lista_puncte = new List<string>();

            DataRelation relation_pt = null;
            if (dt_welds_with_pmc != null && dt_welds_with_pmc.Rows.Count > 0)
            {
                dataset1.Tables.Add(dt_welds_with_pmc);
                dt_welds_with_pmc.TableName = "dt_welds_with_pmc";

                relation_pt = new DataRelation("xxx1", dtwm.Columns[col_pnt], dt_welds_with_pmc.Columns[col_pnt], false);
                dataset1.Relations.Add(relation_pt);
            }

            string status_rock_shield = "END";
            for (int i = 0; i < dtwm.Rows.Count; ++i)
            {
                string feature_rs1 = "xx";
                if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                {
                    feature_rs1 = Convert.ToString(dtwm.Rows[i][col_feat_code]);
                }

                #region null values
                if (dtwm.Rows[i][col_pnt] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_pnt]) == "")
                {
                    dt_errors.Rows.Add();
                    dt_errors.Rows[dt_errors.Rows.Count - 1][1] = feature_rs1;
                    dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_1.Text + Convert.ToString(i + start_row);
                    dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Point ID Specified";
                    ++nr_null_values;
                    lista_puncte.Add("null");
                }

                if (dtwm.Rows[i][col_pnt] != DBNull.Value)
                {
                    string pt1 = Convert.ToString(dtwm.Rows[i][col_pnt]).ToUpper();
                    int index1 = -1;



                    if (dtwm.Rows[i][col_y] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_y]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][1] = feature_rs1;
                        dt_errors.Rows[index1][3] = textBox_2.Text + Convert.ToString(i + start_row);
                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No Northing Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Northing Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }
                        ++nr_null_values;
                    }
                    else
                    {
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_y])) == false)
                        {
                            bool adauga = false;
                            if (lista_puncte.Contains(pt1) == false)
                            {
                                lista_puncte.Add(pt1);
                                adauga = true;
                            }
                            index1 = lista_puncte.IndexOf(pt1);
                            if (adauga == true)
                            {
                                dt_errors.Rows.Add();
                                index1 = dt_errors.Rows.Count - 1;
                            }
                            dt_errors.Rows[index1][0] = pt1;
                            dt_errors.Rows[index1][1] = feature_rs1;
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_y]);
                            dt_errors.Rows[index1][3] = textBox_2.Text + Convert.ToString(i + start_row);
                            if (adauga == true)
                            {
                                dt_errors.Rows[index1][4] = "Northing not Numeric";
                            }
                            else
                            {
                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Northing not Numeric";
                                dt_errors.Rows[index1][4] = Existing_error;
                            }
                            ++nr_null_values;
                        }
                    }
                    if (dtwm.Rows[i][col_x] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_x]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][1] = feature_rs1;
                        dt_errors.Rows[index1][3] = textBox_3.Text + Convert.ToString(i + start_row);
                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No Easting Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Easting Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }

                        ++nr_null_values;
                    }
                    else
                    {
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_x])) == false)
                        {
                            bool adauga = false;
                            if (lista_puncte.Contains(pt1) == false)
                            {
                                lista_puncte.Add(pt1);
                                adauga = true;
                            }
                            index1 = lista_puncte.IndexOf(pt1);
                            if (adauga == true)
                            {
                                dt_errors.Rows.Add();
                                index1 = dt_errors.Rows.Count - 1;
                            }
                            dt_errors.Rows[index1][0] = pt1;
                            dt_errors.Rows[index1][1] = feature_rs1;
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_x]);
                            dt_errors.Rows[index1][3] = textBox_3.Text + Convert.ToString(i + start_row);
                            if (adauga == true)
                            {
                                dt_errors.Rows[index1][4] = "Easting not Numeric";
                            }
                            else
                            {
                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Easting not Numeric";
                                dt_errors.Rows[index1][4] = Existing_error;
                            }

                            ++nr_null_values;
                        }
                    }

                    if (dtwm.Rows[i][col_z] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_z]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][1] = feature_rs1;
                        dt_errors.Rows[index1][3] = textBox_4.Text + Convert.ToString(i + start_row);

                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No Elevation Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Elevation Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;
                        ++nr_null_values;
                    }
                    else
                    {
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_z])) == false)
                        {
                            bool adauga = false;
                            if (lista_puncte.Contains(pt1) == false)
                            {
                                lista_puncte.Add(pt1);
                                adauga = true;
                            }
                            index1 = lista_puncte.IndexOf(pt1);
                            if (adauga == true)
                            {
                                dt_errors.Rows.Add();
                                index1 = dt_errors.Rows.Count - 1;
                            }
                            dt_errors.Rows[index1][0] = pt1;
                            dt_errors.Rows[index1][1] = feature_rs1;
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_z]);
                            dt_errors.Rows[index1][3] = textBox_4.Text + Convert.ToString(i + start_row);
                            if (adauga == true)
                            {
                                dt_errors.Rows[index1][4] = "Elevation not Numeric";
                            }
                            else
                            {
                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Elevation not Numeric";
                                dt_errors.Rows[index1][4] = Existing_error;
                            }

                            string x = "";
                            if (dtwm.Rows[i][col_x] != DBNull.Value)
                            {
                                x = Convert.ToString(dtwm.Rows[i][col_x]);
                            }
                            string y = "";
                            if (dtwm.Rows[i][col_y] != DBNull.Value)
                            {
                                y = Convert.ToString(dtwm.Rows[i][col_y]);
                            }

                            dt_errors.Rows[index1][5] = x;
                            dt_errors.Rows[index1][6] = y;
                            ++nr_null_values;
                        }
                    }
                    if (dtwm.Rows[i][col_feat_code] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_feat_code]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][3] = textBox_5.Text + Convert.ToString(i + start_row);

                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No Feature Code Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Feature Code Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }
                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }

                    if (dtwm.Rows[i][col_descr] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_descr]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][1] = feature_rs1;
                        dt_errors.Rows[index1][3] = textBox_6.Text + Convert.ToString(i + start_row);

                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No Description Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Description Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }

                    if (dtwm.Rows[i][col_sta_lin] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_sta_lin]) == "")
                    {
                        bool adauga = false;
                        if (lista_puncte.Contains(pt1) == false)
                        {
                            lista_puncte.Add(pt1);
                            adauga = true;
                        }
                        index1 = lista_puncte.IndexOf(pt1);
                        if (adauga == true)
                        {
                            dt_errors.Rows.Add();
                            index1 = dt_errors.Rows.Count - 1;
                        }
                        dt_errors.Rows[index1][0] = pt1;
                        dt_errors.Rows[index1][1] = feature_rs1;
                        dt_errors.Rows[index1][3] = textBox_7.Text + Convert.ToString(i + start_row);

                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "No station Linear Specified";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No station Linear Specified";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }

                        string x = "";
                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }
                    else
                    {
                        bool adauga = false;
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_sta_lin])) == false)
                        {
                            if (lista_puncte.Contains(pt1) == false)
                            {
                                lista_puncte.Add(pt1);
                                adauga = true;
                            }
                            index1 = lista_puncte.IndexOf(pt1);
                            if (adauga == true)
                            {
                                dt_errors.Rows.Add();
                                index1 = dt_errors.Rows.Count - 1;
                            }
                            dt_errors.Rows[index1][0] = pt1;
                            dt_errors.Rows[index1][1] = feature_rs1;
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_sta_lin]);
                            dt_errors.Rows[index1][3] = textBox_7.Text + Convert.ToString(i + start_row);

                            if (adauga == true)
                            {
                                dt_errors.Rows[index1][4] = "Station Linear not Numeric";
                            }
                            else
                            {
                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Station Linear not Numeric";
                                dt_errors.Rows[index1][4] = Existing_error;
                            }

                            string x = "";
                            if (dtwm.Rows[i][col_x] != DBNull.Value)
                            {
                                x = Convert.ToString(dtwm.Rows[i][col_x]);
                            }
                            string y = "";
                            if (dtwm.Rows[i][col_y] != DBNull.Value)
                            {
                                y = Convert.ToString(dtwm.Rows[i][col_y]);
                            }

                            dt_errors.Rows[index1][5] = x;
                            dt_errors.Rows[index1][6] = y;

                            ++nr_null_values;
                        }
                    }

                    if (dtwm.Rows[i][col_feat_code] != DBNull.Value)
                    {
                        if (Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "WLD")
                        {
                            bool is_fab1 = false;
                            if (dtwm.Rows[i][col_mm_bk] != DBNull.Value)
                            {
                                if (Convert.ToString(dtwm.Rows[i][col_mm_bk]).ToUpper() == "FAB")
                                {
                                    is_fab1 = true;
                                }
                            }

                            bool is_fab2 = false;
                            if (dtwm.Rows[i][col_mm_ahd] != DBNull.Value)
                            {
                                if (Convert.ToString(dtwm.Rows[i][col_mm_ahd]).ToUpper() == "FAB")
                                {
                                    is_fab2 = true;
                                }
                            }

                            if ((dtwm.Rows[i][col_wall_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_wall_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_11.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Wall Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Wall Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_wall_bk])) == false && is_fab1 == false)
                                {
                                    bool adauga = false;
                                    if (lista_puncte.Contains(pt1) == false)
                                    {
                                        lista_puncte.Add(pt1);
                                        adauga = true;
                                    }
                                    index1 = lista_puncte.IndexOf(pt1);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows.Add();
                                        index1 = dt_errors.Rows.Count - 1;
                                    }
                                    dt_errors.Rows[index1][0] = pt1;
                                    dt_errors.Rows[index1][1] = feature_rs1;
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_wall_bk]);
                                    dt_errors.Rows[index1][3] = textBox_11.Text + Convert.ToString(i + start_row);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows[index1][4] = "Wall Back not Numeric";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Wall Back not Numeric";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }

                                    string x = "";
                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }
                            if ((dtwm.Rows[i][col_mm_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_mm_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_9.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No MMID Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No MMID Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_pipe_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_pipe_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_12.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Pipe ID Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe ID Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_heat_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_heat_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_13.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Heat Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Heat Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;

                            }
                            if ((dtwm.Rows[i][col_coat_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_coat_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_14.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Coating Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Coating Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_grade_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_grade_bk]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_15.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Pipe Grade Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe Grade Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if (dtwm.Rows[i][col_mm_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_mm_ahd]) == "")
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_17.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No MMID Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No MMID Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_wall_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_wall_ahd]) == "") && is_fab1 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Wall Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Wall Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_wall_ahd])) == false && is_fab1 == false)
                                {
                                    bool adauga = false;
                                    if (lista_puncte.Contains(pt1) == false)
                                    {
                                        lista_puncte.Add(pt1);
                                        adauga = true;
                                    }
                                    index1 = lista_puncte.IndexOf(pt1);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows.Add();
                                        index1 = dt_errors.Rows.Count - 1;
                                    }
                                    dt_errors.Rows[index1][0] = pt1;
                                    dt_errors.Rows[index1][1] = feature_rs1;
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_wall_ahd]);
                                    dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);

                                    if (adauga == true)
                                    {
                                        dt_errors.Rows[index1][4] = "Wall Ahead not Numeric";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Wall Ahead not Numeric";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }

                                    string x = "";
                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }
                            if ((dtwm.Rows[i][col_pipe_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_pipe_ahd]) == "") && is_fab2 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_20.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Pipe ID Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe ID Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_heat_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_heat_ahd]) == "") && is_fab2 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_21.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Heat Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Heat Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_coat_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_coat_ahd]) == "") && is_fab2 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_22.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Coating Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Coating Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;


                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col_grade_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_grade_ahd]) == "") && is_fab2 == false)
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_23.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Grade Pipe Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Grade Pipe Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;






                                ++nr_null_values;
                            }


                            if (dtwm.Rows[i][col_manufacture_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_manufacture_bk]) == "")
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_16.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Back manufacturer Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Back manufacturer Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                ++nr_null_values;
                            }
                            if (dtwm.Rows[i][col_manufacture_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_manufacture_ahd]) == "")
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_24.Text + Convert.ToString(i + start_row);
                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Ahead manufacturer Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Ahead manufacturer Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                ++nr_null_values;
                            }


                            if (dtwm.Rows[i][col_length_bk] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_length_bk]) == "")
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_10.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Back length Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Back Length Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                bool adauga = false;
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_length_bk])) == false)
                                {
                                    if (lista_puncte.Contains(pt1) == false)
                                    {
                                        lista_puncte.Add(pt1);
                                        adauga = true;
                                    }
                                    index1 = lista_puncte.IndexOf(pt1);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows.Add();
                                        index1 = dt_errors.Rows.Count - 1;
                                    }
                                    dt_errors.Rows[index1][0] = pt1;
                                    dt_errors.Rows[index1][1] = feature_rs1;
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_length_bk]);
                                    dt_errors.Rows[index1][3] = textBox_10.Text + Convert.ToString(i + start_row);

                                    if (adauga == true)
                                    {
                                        dt_errors.Rows[index1][4] = "Back Length not Numeric";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Back length not Numeric";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }

                                    string x = "";
                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }

                            if (dtwm.Rows[i][col_length_ahd] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col_length_ahd]) == "")
                            {
                                bool adauga = false;
                                if (lista_puncte.Contains(pt1) == false)
                                {
                                    lista_puncte.Add(pt1);
                                    adauga = true;
                                }
                                index1 = lista_puncte.IndexOf(pt1);
                                if (adauga == true)
                                {
                                    dt_errors.Rows.Add();
                                    index1 = dt_errors.Rows.Count - 1;
                                }
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature_rs1;
                                dt_errors.Rows[index1][3] = textBox_18.Text + Convert.ToString(i + start_row);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "No Length ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Length ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col_x] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col_x]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col_y] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col_y]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                bool adauga = false;
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col_length_ahd])) == false)
                                {
                                    if (lista_puncte.Contains(pt1) == false)
                                    {
                                        lista_puncte.Add(pt1);
                                        adauga = true;
                                    }
                                    index1 = lista_puncte.IndexOf(pt1);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows.Add();
                                        index1 = dt_errors.Rows.Count - 1;
                                    }
                                    dt_errors.Rows[index1][0] = pt1;
                                    dt_errors.Rows[index1][1] = feature_rs1;
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col_length_ahd]);
                                    dt_errors.Rows[index1][3] = textBox_18.Text + Convert.ToString(i + start_row);

                                    if (adauga == true)
                                    {
                                        dt_errors.Rows[index1][4] = "Length ahead not Numeric";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Length ahead not Numeric";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }

                                    string x = "";
                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }


                        }
                    }
                }
                #endregion


                if (feature_rs1 != "xx")
                {
                    if (dtwm.Rows[i][col_pnt] != DBNull.Value)
                    {
                        string pt1 = Convert.ToString(dtwm.Rows[i][col_pnt]).ToUpper();
                        #region BEND
                        if (Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "BEND" || Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "ELBOW")
                        {
                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_pnt] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_mm_ahd] = dtwm.Rows[i][col_mm_ahd];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "B";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion

                        #region loose end
                        if (Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "LOOSE_END" || Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "LE")
                        {
                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_pnt] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_mm_ahd] = dtwm.Rows[i][col_mm_ahd];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "X";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion

                        #region WELD
                        if (Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[i][col_feat_code]).ToUpper() == "WLD")
                        {

                            dt_welds.Rows.Add();
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_pnt] = pt1;
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_sta_lin] = dtwm.Rows[i][col_sta_lin];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_wall_bk] = dtwm.Rows[i][col_wall_bk];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_wall_ahd] = dtwm.Rows[i][col_wall_ahd];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_coat_bk] = dtwm.Rows[i][col_coat_bk];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col_coat_ahd] = dtwm.Rows[i][col_coat_ahd];
                            dt_welds.Rows[dt_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);

                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_pnt] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_y] = dtwm.Rows[i][col_y];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_x] = dtwm.Rows[i][col_x];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_z] = dtwm.Rows[i][col_z];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col_mm_ahd] = dtwm.Rows[i][col_mm_ahd];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "W";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_17.Text + Convert.ToString(i + start_row);

                            bool is_loose_end = false;

                            if (dtwm.Rows[i].GetChildRows(relation_pt).Length > 0)
                            {
                                if (dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"] != DBNull.Value && Convert.ToString(dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"]) == "LE AHEAD")
                                {
                                    is_loose_end = true;
                                }
                            }

                            if (is_loose_end == false)
                            {
                                int index1 = -1;

                                #region mmid back-ahead
                                if (dtwm.Rows[i][col_mm_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_mm_ahd]) != "")
                                {
                                    string MM1 = Convert.ToString(dtwm.Rows[i][col_mm_ahd]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col_mm_bk] != DBNull.Value)
                                                {
                                                    string MM2 = Convert.ToString(dtwm.Rows[j][col_mm_bk]);
                                                    if (MM1.ToUpper() != MM2.ToUpper())
                                                    {
                                                        bool adauga = false;
                                                        if (lista_puncte.Contains(pt1) == false)
                                                        {
                                                            lista_puncte.Add(pt1);
                                                            adauga = true;
                                                        }
                                                        index1 = lista_puncte.IndexOf(pt1);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            index1 = dt_errors.Rows.Count - 1;
                                                        }
                                                        dt_errors.Rows[index1][0] = pt1;
                                                        dt_errors.Rows[index1][1] = feature_rs1;
                                                        dt_errors.Rows[index1][2] = "MMID: " + MM1 + " vs. " + MM2;
                                                        dt_errors.Rows[index1][3] = textBox_17.Text + Convert.ToString(i + start_row);

                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "MM id Ahead vs next row id Back mismatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "MM id Ahead vs next row id Back mismatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }

                                                        string x = "";
                                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                                        }

                                                        dt_errors.Rows[index1][5] = x;
                                                        dt_errors.Rows[index1][6] = y;
                                                    }
                                                }
                                                j = dtwm.Rows.Count;
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region wall back-ahead
                                if (dtwm.Rows[i][col_wall_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_wall_ahd]) != "" &&
                                    dtwm.Rows[i][col_wall_bk] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_wall_bk]) != "")
                                {
                                    bool raise_error = false;
                                    string wall1_ahead = Convert.ToString(dtwm.Rows[i][col_wall_ahd]);
                                    string wall1_back = Convert.ToString(dtwm.Rows[i][col_wall_bk]);
                                    string pt_pmc = "";

                                    if (dtwm.Rows[i].GetChildRows(relation_pt).Length > 0)
                                    {
                                        if (dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"] != DBNull.Value)
                                        {
                                            pt_pmc = Convert.ToString(dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"]);
                                        }
                                    }

                                    string err_back = wall1_back;
                                    string err_ahead = wall1_ahead;

                                    if (wall1_ahead.ToUpper() != wall1_back.ToUpper())
                                    {
                                        if (pt_pmc == "")
                                        {
                                            raise_error = true;
                                        }
                                        else
                                        {
                                            string descr1 = "";
                                            for (int k = 0; k < dtwm.Rows.Count; ++k)
                                            {
                                                if (dtwm.Rows[k][col_pnt] != DBNull.Value)
                                                {
                                                    string pt_found = Convert.ToString(dtwm.Rows[k][col_pnt]);
                                                    if (pt_found == pt_pmc)
                                                    {
                                                        if (dtwm.Rows[k][col_descr] != DBNull.Value)
                                                        {
                                                            descr1 = Convert.ToString(dtwm.Rows[k][col_descr]);
                                                        }
                                                        k = dtwm.Rows.Count;
                                                    }
                                                }
                                            }
                                            if (descr1 != "")
                                            {
                                                if (descr1.Contains("TO " + wall1_ahead) == false || descr1.Contains(wall1_back) == false)
                                                {
                                                    raise_error = true;
                                                }
                                            }
                                        }
                                    }


                                    bool raise_error1 = false;
                                    string err_back1 = wall1_ahead;
                                    string err_ahead1 = "";

                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col_wall_bk] != DBNull.Value)
                                                {
                                                    string wall2_back = Convert.ToString(dtwm.Rows[j][col_wall_bk]);

                                                    if (wall1_ahead.ToUpper() != wall2_back.ToUpper())
                                                    {
                                                        raise_error1 = true;
                                                        err_ahead1 = wall2_back;
                                                    }
                                                }

                                                j = dtwm.Rows.Count; // as soon as you find another weld you stop going forward
                                            }
                                        }

                                    }

                                    if (raise_error == true)
                                    {
                                        bool adauga = false;
                                        if (lista_puncte.Contains(pt1) == false)
                                        {
                                            lista_puncte.Add(pt1);
                                            adauga = true;
                                        }
                                        index1 = lista_puncte.IndexOf(pt1);
                                        if (adauga == true)
                                        {
                                            dt_errors.Rows.Add();
                                            index1 = dt_errors.Rows.Count - 1;
                                        }

                                        dt_errors.Rows[index1][0] = pt1;
                                        dt_errors.Rows[index1][1] = feature_rs1;
                                        dt_errors.Rows[index1][2] = err_back + " vs. " + err_ahead;
                                        dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);
                                        if (adauga == true)
                                        {
                                            dt_errors.Rows[index1][4] = "PMC point missing";
                                        }
                                        else
                                        {
                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "PMC point missing";
                                            dt_errors.Rows[index1][4] = Existing_error;
                                        }
                                        string x = "";
                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                        {
                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                        }
                                        string y = "";
                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                        {
                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                        }
                                        dt_errors.Rows[index1][5] = x;
                                        dt_errors.Rows[index1][6] = y;
                                    }

                                    if (raise_error1 == true)
                                    {
                                        bool adauga = false;
                                        if (lista_puncte.Contains(pt1) == false)
                                        {
                                            lista_puncte.Add(pt1);
                                            adauga = true;
                                        }
                                        index1 = lista_puncte.IndexOf(pt1);
                                        if (adauga == true)
                                        {
                                            dt_errors.Rows.Add();
                                            index1 = dt_errors.Rows.Count - 1;
                                        }

                                        dt_errors.Rows[index1][0] = pt1;
                                        dt_errors.Rows[index1][1] = feature_rs1;
                                        dt_errors.Rows[index1][2] = err_back1 + " vs. " + err_ahead1;
                                        dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);
                                        if (adauga == true)
                                        {
                                            dt_errors.Rows[index1][4] = "WT Ahead vs next row WT Back mismatch";
                                        }
                                        else
                                        {
                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "WT Ahead vs next row WT Back mismatch";
                                            dt_errors.Rows[index1][4] = Existing_error;
                                        }
                                        string x = "";
                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                        {
                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                        }
                                        string y = "";
                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                        {
                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                        }
                                        dt_errors.Rows[index1][5] = x;
                                        dt_errors.Rows[index1][6] = y;
                                    }
                                    else
                                    {
                                        if (Wgen_main_form.use_pmc_as_cc == true && raise_error == false)
                                        {
                                            if (dtwm.Rows[i][col_coat_ahd] != DBNull.Value &&
                                                Convert.ToString(dtwm.Rows[i][col_coat_ahd]) != "" &&
                                                dtwm.Rows[i][col_coat_bk] != DBNull.Value &&
                                                Convert.ToString(dtwm.Rows[i][col_coat_bk]) != "")
                                            {

                                                string coating1_ahead = Convert.ToString(dtwm.Rows[i][col_coat_ahd]);
                                                string coating1_back = Convert.ToString(dtwm.Rows[i][col_coat_bk]);



                                                bool raise_error_cc = false;
                                                string val_back1 = coating1_ahead;
                                                string val_ahead1 = "";

                                                if (i < dtwm.Rows.Count - 1)
                                                {
                                                    for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                                    {
                                                        if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                                        {
                                                            if (dtwm.Rows[j][col_coat_bk] != DBNull.Value)
                                                            {
                                                                string coating2_back = Convert.ToString(dtwm.Rows[j][col_coat_bk]);

                                                                if (coating1_ahead.ToUpper() != coating2_back.ToUpper())
                                                                {
                                                                    raise_error_cc = true;
                                                                    val_ahead1 = coating2_back;
                                                                }
                                                            }

                                                            j = dtwm.Rows.Count; // as soon as you find another weld you stop going forward
                                                        }
                                                    }

                                                }

                                                if (raise_error_cc == true)
                                                {
                                                    bool adauga = false;
                                                    if (lista_puncte.Contains(pt1) == false)
                                                    {
                                                        lista_puncte.Add(pt1);
                                                        adauga = true;
                                                    }
                                                    index1 = lista_puncte.IndexOf(pt1);
                                                    if (adauga == true)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        index1 = dt_errors.Rows.Count - 1;
                                                    }

                                                    dt_errors.Rows[index1][0] = pt1;
                                                    dt_errors.Rows[index1][1] = feature_rs1;
                                                    dt_errors.Rows[index1][2] = val_back1 + " vs. " + val_ahead1;
                                                    dt_errors.Rows[index1][3] = textBox_22.Text + Convert.ToString(i + start_row);
                                                    if (adauga == true)
                                                    {
                                                        dt_errors.Rows[index1][4] = "Coating Ahead vs next row Coating Back mismatch";
                                                    }
                                                    else
                                                    {
                                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating Ahead vs next row Coating Back mismatch";
                                                        dt_errors.Rows[index1][4] = Existing_error;
                                                    }
                                                    string x = "";
                                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                                    }
                                                    string y = "";
                                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                                    }
                                                    dt_errors.Rows[index1][5] = x;
                                                    dt_errors.Rows[index1][6] = y;
                                                }


                                            }
                                        }
                                    }


                                }




                                #endregion

                                #region pipe back-ahead
                                if (dtwm.Rows[i][col_pipe_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_pipe_ahd]) != "")
                                {
                                    string pipeid1 = Convert.ToString(dtwm.Rows[i][col_pipe_ahd]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col_pipe_bk] != DBNull.Value)
                                                {
                                                    string pipeid2 = Convert.ToString(dtwm.Rows[j][col_pipe_bk]);

                                                    if (pipeid1.ToUpper() != pipeid2.ToUpper())
                                                    {
                                                        bool adauga = false;
                                                        if (lista_puncte.Contains(pt1) == false)
                                                        {
                                                            lista_puncte.Add(pt1);
                                                            adauga = true;
                                                        }
                                                        index1 = lista_puncte.IndexOf(pt1);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            index1 = dt_errors.Rows.Count - 1;
                                                        }

                                                        dt_errors.Rows[index1][0] = pt1;
                                                        dt_errors.Rows[index1][1] = feature_rs1;
                                                        dt_errors.Rows[index1][2] = "PipeID: " + pipeid1 + " vs. " + pipeid2;
                                                        dt_errors.Rows[index1][3] = textBox_20.Text + Convert.ToString(i + start_row);

                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Pipe id Ahead vs next row Back mismatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Pipe id Ahead vs next row Back mismatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }

                                                        string x = "";
                                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                                        }

                                                        dt_errors.Rows[index1][5] = x;
                                                        dt_errors.Rows[index1][6] = y;

                                                    }


                                                }

                                                j = dtwm.Rows.Count;
                                            }

                                        }
                                    }
                                }
                                #endregion

                                #region heat back-ahead
                                if (dtwm.Rows[i][col_heat_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_heat_ahd]) != "")
                                {
                                    string heat1 = Convert.ToString(dtwm.Rows[i][col_heat_ahd]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col_heat_bk] != DBNull.Value)
                                                {
                                                    string heat2 = Convert.ToString(dtwm.Rows[j][col_heat_bk]);
                                                    if (heat1.ToUpper() != heat2.ToUpper())
                                                    {
                                                        bool adauga = false;
                                                        if (lista_puncte.Contains(pt1) == false)
                                                        {
                                                            lista_puncte.Add(pt1);
                                                            adauga = true;
                                                        }
                                                        index1 = lista_puncte.IndexOf(pt1);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            index1 = dt_errors.Rows.Count - 1;
                                                        }
                                                        dt_errors.Rows[index1][0] = pt1;
                                                        dt_errors.Rows[index1][1] = feature_rs1;
                                                        dt_errors.Rows[index1][2] = "Heat: " + heat1 + " vs. " + heat2;
                                                        dt_errors.Rows[index1][3] = textBox_21.Text + Convert.ToString(i + start_row);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Heat# Ahead vs next row Back mismatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Heat# Ahead vs next row Back mismatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }
                                                        string x = "";
                                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                                        }

                                                        dt_errors.Rows[index1][5] = x;
                                                        dt_errors.Rows[index1][6] = y;
                                                    }
                                                }
                                                j = dtwm.Rows.Count;
                                            }
                                        }
                                    }
                                }
                                #endregion

                                if (Wgen_main_form.use_pmc_as_cc == false)
                                {
                                    #region coating back-ahead
                                    if (dtwm.Rows[i][col_coat_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_coat_ahd]) != "" &&
                                        dtwm.Rows[i][col_coat_bk] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_coat_bk]) != "")
                                    {
                                        bool raise_error = false;
                                        string coating1_ahead = Convert.ToString(dtwm.Rows[i][col_coat_ahd]);
                                        string coating1_back = Convert.ToString(dtwm.Rows[i][col_coat_bk]);
                                        string pt_cc = "";

                                        if (dtwm.Rows[i].GetChildRows(relation_pt).Length > 0)
                                        {
                                            if (dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"] != DBNull.Value)
                                            {
                                                pt_cc = Convert.ToString(dtwm.Rows[i].GetChildRows(relation_pt)[0]["PMC"]);
                                            }
                                            if (dtwm.Rows[i].GetChildRows(relation_pt)[0]["CC"] != DBNull.Value)
                                            {
                                                pt_cc = Convert.ToString(dtwm.Rows[i].GetChildRows(relation_pt)[0]["CC"]);
                                            }
                                        }

                                        string val_back = coating1_back;
                                        string val_ahead = coating1_ahead;

                                        if (coating1_ahead.ToUpper() != coating1_back.ToUpper())
                                        {
                                            if (pt_cc == "")
                                            {
                                                raise_error = true;
                                            }
                                            else
                                            {
                                                string descr1 = "";
                                                for (int k = 0; k < dtwm.Rows.Count; ++k)
                                                {
                                                    if (dtwm.Rows[k][col_pnt] != DBNull.Value)
                                                    {
                                                        string pt_found = Convert.ToString(dtwm.Rows[k][col_pnt]);
                                                        if (pt_found == pt_cc)
                                                        {
                                                            if (dtwm.Rows[k][col_descr] != DBNull.Value)
                                                            {
                                                                descr1 = Convert.ToString(dtwm.Rows[k][col_descr]);
                                                            }
                                                            k = dtwm.Rows.Count;
                                                        }
                                                    }
                                                }
                                                if (descr1 != "")
                                                {
                                                    if (descr1.Contains(coating1_back + " TO " + coating1_ahead) == false)
                                                    {
                                                        raise_error = true;
                                                    }
                                                }
                                            }
                                        }


                                        bool raise_error1 = false;
                                        string val_back1 = coating1_ahead;
                                        string val_ahead1 = "";

                                        if (i < dtwm.Rows.Count - 1)
                                        {
                                            for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                            {
                                                if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                                {
                                                    if (dtwm.Rows[j][col_coat_bk] != DBNull.Value)
                                                    {
                                                        string coating2_back = Convert.ToString(dtwm.Rows[j][col_coat_bk]);

                                                        if (coating1_ahead.ToUpper() != coating2_back.ToUpper())
                                                        {
                                                            raise_error1 = true;
                                                            val_ahead1 = coating2_back;
                                                        }
                                                    }

                                                    j = dtwm.Rows.Count; // as soon as you find another weld you stop going forward
                                                }
                                            }

                                        }

                                        if (raise_error == true)
                                        {
                                            bool adauga = false;
                                            if (lista_puncte.Contains(pt1) == false)
                                            {
                                                lista_puncte.Add(pt1);
                                                adauga = true;
                                            }
                                            index1 = lista_puncte.IndexOf(pt1);
                                            if (adauga == true)
                                            {
                                                dt_errors.Rows.Add();
                                                index1 = dt_errors.Rows.Count - 1;
                                            }

                                            dt_errors.Rows[index1][0] = pt1;
                                            dt_errors.Rows[index1][1] = feature_rs1;
                                            dt_errors.Rows[index1][2] = val_back + " vs. " + val_ahead;
                                            dt_errors.Rows[index1][3] = textBox_22.Text + Convert.ToString(i + start_row);
                                            if (adauga == true)
                                            {
                                                dt_errors.Rows[index1][4] = "Coating Change point missing";
                                            }
                                            else
                                            {
                                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating Change point missing";
                                                dt_errors.Rows[index1][4] = Existing_error;
                                            }
                                            string x = "";
                                            if (dtwm.Rows[i][col_x] != DBNull.Value)
                                            {
                                                x = Convert.ToString(dtwm.Rows[i][col_x]);
                                            }
                                            string y = "";
                                            if (dtwm.Rows[i][col_y] != DBNull.Value)
                                            {
                                                y = Convert.ToString(dtwm.Rows[i][col_y]);
                                            }
                                            dt_errors.Rows[index1][5] = x;
                                            dt_errors.Rows[index1][6] = y;
                                        }

                                        if (raise_error1 == true)
                                        {
                                            bool adauga = false;
                                            if (lista_puncte.Contains(pt1) == false)
                                            {
                                                lista_puncte.Add(pt1);
                                                adauga = true;
                                            }
                                            index1 = lista_puncte.IndexOf(pt1);
                                            if (adauga == true)
                                            {
                                                dt_errors.Rows.Add();
                                                index1 = dt_errors.Rows.Count - 1;
                                            }

                                            dt_errors.Rows[index1][0] = pt1;
                                            dt_errors.Rows[index1][1] = feature_rs1;
                                            dt_errors.Rows[index1][2] = val_back1 + " vs. " + val_ahead1;
                                            dt_errors.Rows[index1][3] = textBox_22.Text + Convert.ToString(i + start_row);
                                            if (adauga == true)
                                            {
                                                dt_errors.Rows[index1][4] = "Coating Ahead vs next row Coating Back mismatch";
                                            }
                                            else
                                            {
                                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating Ahead vs next row Coating Back mismatch";
                                                dt_errors.Rows[index1][4] = Existing_error;
                                            }
                                            string x = "";
                                            if (dtwm.Rows[i][col_x] != DBNull.Value)
                                            {
                                                x = Convert.ToString(dtwm.Rows[i][col_x]);
                                            }
                                            string y = "";
                                            if (dtwm.Rows[i][col_y] != DBNull.Value)
                                            {
                                                y = Convert.ToString(dtwm.Rows[i][col_y]);
                                            }
                                            dt_errors.Rows[index1][5] = x;
                                            dt_errors.Rows[index1][6] = y;
                                        }


                                    }
                                    #endregion
                                }



                                #region grade back-ahead
                                if (dtwm.Rows[i][col_grade_ahd] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_grade_ahd]) != "")
                                {
                                    string grade1 = Convert.ToString(dtwm.Rows[i][col_grade_ahd]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col_feat_code]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col_grade_bk] != DBNull.Value)
                                                {
                                                    string grade2 = Convert.ToString(dtwm.Rows[j][col_grade_bk]);
                                                    if (grade1.ToUpper() != grade2.ToUpper())
                                                    {
                                                        bool adauga = false;
                                                        if (lista_puncte.Contains(pt1) == false)
                                                        {
                                                            lista_puncte.Add(pt1);
                                                            adauga = true;
                                                        }
                                                        index1 = lista_puncte.IndexOf(pt1);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows.Add();
                                                            index1 = dt_errors.Rows.Count - 1;
                                                        }
                                                        dt_errors.Rows[index1][0] = pt1;
                                                        dt_errors.Rows[index1][1] = feature_rs1;
                                                        dt_errors.Rows[index1][2] = "Pipe Grade: " + grade1 + " vs. " + grade2;
                                                        dt_errors.Rows[index1][3] = textBox_23.Text + Convert.ToString(i + start_row);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Pipe Grade Ahead vs next row Back mismatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Pipe Grade Ahead vs next row Back mismatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }
                                                        string x = "";
                                                        if (dtwm.Rows[i][col_x] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col_x]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col_y] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col_y]);
                                                        }

                                                        dt_errors.Rows[index1][5] = x;
                                                        dt_errors.Rows[index1][6] = y;
                                                    }
                                                }

                                                j = dtwm.Rows.Count;
                                            }
                                        }
                                    }
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region ROCK SHIELD
                        if (feature_rs1.ToUpper() == "ROCK_SHIELD")
                        {
                            if (dtwm.Rows[i][col_descr] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col_descr]) != "" && dtwm.Rows[i][col_y] != DBNull.Value && dtwm.Rows[i][col_x] != DBNull.Value)
                            {
                                string descr = Convert.ToString(dtwm.Rows[i][col_descr]);
                                if (descr.ToUpper() == status_rock_shield || (descr.ToUpper() != "BEGIN" && descr.ToUpper() != "END"))
                                {

                                    int index1 = -1;

                                    bool adauga = false;
                                    if (lista_puncte.Contains(pt1) == false)
                                    {
                                        lista_puncte.Add(pt1);
                                        adauga = true;
                                    }
                                    index1 = lista_puncte.IndexOf(pt1);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows.Add();
                                        index1 = dt_errors.Rows.Count - 1;
                                    }
                                    dt_errors.Rows[index1][0] = pt1;
                                    dt_errors.Rows[index1][1] = feature_rs1;
                                    dt_errors.Rows[index1][2] = "Rock Shield: " + descr + " repetition";
                                    dt_errors.Rows[index1][3] = textBox_5.Text + Convert.ToString(i + start_row);
                                    if (adauga == true)
                                    {
                                        dt_errors.Rows[index1][4] = "Rock Shield Start/End mismatch";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Rock Shield Start/End mismatch";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }
                                    string x = "";
                                    if (dtwm.Rows[i][col_x] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col_x]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col_y] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col_y]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;
                                }
                                else
                                {
                                    if (status_rock_shield == "BEGIN")
                                    {
                                        status_rock_shield = "END";
                                    }
                                    else
                                    {
                                        status_rock_shield = "BEGIN";
                                    }
                                }
                            }


                        }
                        #endregion
                    }
                }
            }

            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_errors);


            if (dt_bend_welds.Rows.Count > 0)
            {
                #region lengths checks
                if (Wgen_main_form.dt_ground_tally != null && Wgen_main_form.dt_ground_tally.Rows.Count > 0)
                {
                    dataset1.Tables.Add(dt_bend_welds);
                    dataset1.Tables.Add(Wgen_main_form.dt_ground_tally);

                    DataRelation relation_pt1 = new DataRelation("xxx8", dt_bend_welds.Columns[col_mm_ahd], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
                    dataset1.Relations.Add(relation_pt1);

                    double x0 = -1;
                    double y0 = -1;
                    double z0 = -1;
                    double d1 = 0;

                    int prev_i = 0;

                    bool is_loose_end = false;

                    //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_bend_welds);

                    for (int i = 0; i < dt_bend_welds.Rows.Count - 1; ++i)
                    {
                        if (dt_bend_welds.Rows[i][col_pnt] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col_y] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col_x] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col_z] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col_pnt] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col_y] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col_x] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col_z] != DBNull.Value)
                        {
                            double x1 = Convert.ToDouble(dt_bend_welds.Rows[i][col_x]);
                            double y1 = Convert.ToDouble(dt_bend_welds.Rows[i][col_y]);
                            double z1 = Convert.ToDouble(dt_bend_welds.Rows[i][col_z]);
                            string pt1 = Convert.ToString(dt_bend_welds.Rows[i][col_pnt]);
                            string ft1 = Convert.ToString(dt_bend_welds.Rows[i]["feature_code"]);
                            if (ft1 == "X")
                            {
                                is_loose_end = true;
                            }

                            if (i > 0)
                            {

                                d1 = d1 + Math.Pow(Math.Pow(x0 - x1, 2) + Math.Pow(y0 - y1, 2) + Math.Pow(z0 - z1, 2), 0.5);

                                if (ft1 == "W")
                                {
                                    if (is_loose_end == false)
                                    {
                                        int nr_total = dt_bend_welds.Rows[prev_i].GetChildRows(relation_pt1).Length;
                                        if (nr_total > 0)
                                        {
                                            double d2 = -1;
                                            if (dt_bend_welds.Rows[prev_i].GetChildRows(relation_pt1)[0][colgt4] != DBNull.Value)
                                            {
                                                d2 = Convert.ToDouble(dt_bend_welds.Rows[prev_i].GetChildRows(relation_pt1)[0][colgt4]);
                                            }
                                            if (dt_bend_welds.Rows[prev_i].GetChildRows(relation_pt1)[0][colgt5] != DBNull.Value)
                                            {
                                                d2 = Convert.ToDouble(dt_bend_welds.Rows[prev_i].GetChildRows(relation_pt1)[0][colgt5]);
                                            }
                                            if (d2 >= 0)
                                            {
                                                if (Math.Abs(d2 - d1) >= length_check_tolerance)
                                                {
                                                    bool adauga = false;
                                                    if (lista_puncte.Contains(pt1) == false)
                                                    {
                                                        lista_puncte.Add(pt1);
                                                        adauga = true;
                                                    }
                                                    int index1 = lista_puncte.IndexOf(pt1);
                                                    if (adauga == true)
                                                    {
                                                        dt_errors.Rows.Add();
                                                        index1 = dt_errors.Rows.Count - 1;
                                                    }

                                                    string pt2 = Convert.ToString(dt_bend_welds.Rows[prev_i][col_pnt]);
                                                    string MM_ahead = "line 4794 - wm_checks";
                                                    if (dt_bend_welds.Rows[prev_i][col_mm_ahd] != DBNull.Value)
                                                    {
                                                        MM_ahead = Convert.ToString(dt_bend_welds.Rows[prev_i][col_mm_ahd]);
                                                    }


                                                    dt_errors.Rows[index1][0] = pt2 + "-" + pt1;
                                                    dt_errors.Rows[index1][1] = "WELD TO WELD";
                                                    dt_errors.Rows[index1][2] = "MM: " + MM_ahead + " - Ground Tally: " + Convert.ToString(Math.Round(d2, 2)) + " vs. calc: " + Convert.ToString(Math.Round(d1, 2));
                                                    dt_errors.Rows[index1][3] = Convert.ToString(dt_bend_welds.Rows[prev_i]["address"]);


                                                    if (adauga == true)
                                                    {
                                                        dt_errors.Rows[index1][4] = "Length mismatch with ground tally";
                                                    }
                                                    else
                                                    {
                                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Length mismatch with ground tally";
                                                        dt_errors.Rows[index1][4] = Existing_error;
                                                    }

                                                    string x = "";
                                                    if (dt_bend_welds.Rows[prev_i][col_x] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(dt_bend_welds.Rows[prev_i][col_x]);
                                                    }
                                                    string y = "";
                                                    if (dt_bend_welds.Rows[prev_i][col_y] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(dt_bend_welds.Rows[prev_i][col_y]);
                                                    }
                                                    dt_errors.Rows[index1][5] = x;
                                                    dt_errors.Rows[index1][6] = y;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        is_loose_end = false;
                                    }

                                    d1 = 0;
                                    prev_i = i;
                                }

                                if (ft1 == "B")
                                {
                                    prev_i = i;
                                }

                            }
                            x0 = x1;
                            y0 = y1;
                            z0 = z1;
                        }
                    }
                    dataset1.Relations.Remove(relation_pt1);
                    dataset1.Tables.Remove(dt_bend_welds);
                    dataset1.Tables.Remove(Wgen_main_form.dt_ground_tally);
                }
                #endregion
            }

            dt_welds = null;

            if (relation_pt != null)
            {
                dataset1.Relations.Remove(relation_pt);
                dataset1.Tables.Remove(dt_welds_with_pmc);
            }

            dataset1.Tables.Remove(dtwm);




            dt_errors = Functions.Sort_data_table(dt_errors, "Error type");


            dt_display = dt_errors.Copy();
            display_errors(dt_display);

            textBox_PM_no_rows.Text = Convert.ToString(dtwm.Rows.Count);
            textBox_PM_no_duplicates.Text = Convert.ToString(nr_duplicates);
            textBox_WM_no_null.Text = Convert.ToString(nr_null_values);
            button_wm_l.Visible = true;
            button_wm_nl.Visible = false;

        }

        private void button_check_weld_map_Click(object sender, EventArgs e)
        {
            Wgen_main_form.dt_weld_map = Functions.Creaza_weldmap_datatable_structure();
            make_first_line_invisible();
            if (comboBox_ws1.Text != "")
            {
                string string1 = comboBox_ws1.Text;
                if (string1.Contains("[") == true && string1.Contains("]") == true)
                {
                    string filename = string1.Substring(string1.IndexOf("]") + 4, string1.Length - (string1.IndexOf("]") + 4));
                    string sheet_name = string1.Substring(1, string1.IndexOf("]") - 1);
                    if (filename.Length > 0 && sheet_name.Length > 0)
                    {
                        set_enable_false();
                        W2 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                        W3 = Functions.Get_opened_worksheet_from_Excel_by_name(Wgen_main_form.tpage_allpts.filename, dismiss_errors_tab);
                        if (W2 != null)
                        {
                            Wgen_main_form.dt_weld_map = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_weld_map, W2, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, textBox_6.Text, textBox_7.Text, textBox_8.Text, textBox_9.Text, textBox_11.Text, textBox_12.Text,
                                 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", true);

                            if (W3 != null)
                            {
                                dt_dismissed_errors = new System.Data.DataTable();
                                dt_dismissed_errors.Columns.Add("Point", typeof(string));
                                dt_dismissed_errors.Columns.Add("Feature Code", typeof(string));
                                dt_dismissed_errors.Columns.Add("Value", typeof(string));
                                dt_dismissed_errors.Columns.Add("Error", typeof(string));

                                dt_dismissed_errors = Functions.Populate_data_table_from_excel(dt_dismissed_errors, W3, start_row,
                                    "A", "B", "C", "D", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", false);
                                if (dt_dismissed_errors.Rows.Count == 0) dt_dismissed_errors = null;
                            }
                            else
                            {
                                dt_dismissed_errors = null;
                            }


                            if (Wgen_main_form.dt_weld_map.Rows.Count > 0)
                            {
                                Wgen_main_form.tpage_weldmap.Hide();
                                Wgen_main_form.tpage_blank.Show();
                                Wgen_main_form.tpage_pipe_manifest.Hide();
                                Wgen_main_form.tpage_pipe_tally.Hide();
                                Wgen_main_form.tpage_allpts.Hide();
                                Wgen_main_form.tpage_build_pipe_tally.Hide();
                                Wgen_main_form.tpage_duplicates.Hide();
                                Wgen_main_form.tpage_dismiss_errors.Hide();
                                Wgen_main_form.tpage_blank.get_label_wait_visible(true);

                                this.Refresh();


                                wm_checks(Wgen_main_form.dt_weld_map);


                            }
                            else
                            {
                                button_wm_l.Visible = false;
                                button_wm_nl.Visible = true;
                            }
                        }
                        set_enable_true();
                    }
                    else
                    {
                        button_wm_l.Visible = false;
                        button_wm_nl.Visible = true;
                    }
                }
                else
                {
                    button_wm_l.Visible = false;
                    button_wm_nl.Visible = true;
                }
            }
            else
            {
                button_wm_l.Visible = false;
                button_wm_nl.Visible = true;
            }
            Wgen_main_form.tpage_weldmap.Show();
            Wgen_main_form.tpage_blank.Hide();
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_dismiss_errors.Hide();
            Wgen_main_form.tpage_blank.get_label_wait_visible(false);

            //Functions.Transfer_datatable_to_new_excel_spreadsheet(Wgen_main_form.dt_weld_map);
        }



        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_named(dt_display, "WeldMapErrors");
        }

        private void DataGridView_error_weld_map_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView_error_weld_map.CurrentCell = dataGridView_error_weld_map.Rows[e.RowIndex].Cells[0];
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

            int index1 = dataGridView_error_weld_map.CurrentCell.RowIndex;
            try
            {
                if (dt_errors.Rows.Count - 1 >= index1)
                {
                    if (W2 != null)
                    {
                        if (dt_errors.Rows[index1]["Excel address"] != DBNull.Value)
                        {
                            string adresa = Convert.ToString(dt_errors.Rows[index1]["Excel address"]);
                            W2.Activate();
                            W2.Range[adresa].Select();
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

            int index1 = dataGridView_error_weld_map.CurrentCell.RowIndex;
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

        public void radioButton_weldmap_CheckChanged(RadioButton radioButton_enlarged)
        {

            System.Drawing.Font regularfont = new Font("Arial", 8.2f, FontStyle.Bold);


            Font englargedFont = new Font("Arial", 10f, FontStyle.Bold);


            Font regularHeader = new Font("Arial", 10f, FontStyle.Bold);

            Font englargedHeader = new Font("Arial", 12f, FontStyle.Bold);
            if (radioButton_enlarged.Checked == true)
            {
                panel7.Location = new System.Drawing.Point(3, 3);
                panel7.Size = new Size(723, 28);
            }
            else
            {
                panel7.Location = new System.Drawing.Point(3, 3);
                panel7.Size = new Size(723, 25);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label12.Location = new System.Drawing.Point(5, 5);
                label12.Size = new Size(93, 23);

                label12.Font = englargedHeader;
                label12.Font = englargedHeader;
            }
            else
            {
                label12.Location = new System.Drawing.Point(3, 3);
                label12.Size = new Size(75, 18);

                label12.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_create_weldmap.Location = new System.Drawing.Point(3, 35);
                button_create_weldmap.Size = new Size(168, 32);

                button_create_weldmap.Font = englargedFont;
            }
            else
            {
                button_create_weldmap.Location = new System.Drawing.Point(3, 31);
                button_create_weldmap.Size = new Size(162, 28);

                button_create_weldmap.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_pipe_manifest.Location = new System.Drawing.Point(3, 68);
                panel_pipe_manifest.Size = new Size(723, 35);
            }
            else
            {
                panel_pipe_manifest.Location = new System.Drawing.Point(3, 63);
                panel_pipe_manifest.Size = new Size(723, 33);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_wm_l.Location = new System.Drawing.Point(695, 3);
                button_wm_l.Size = new Size(24, 24);
            }
            else
            {
                button_wm_l.Location = new System.Drawing.Point(697, 5);
                button_wm_l.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_wm_nl.Location = new System.Drawing.Point(695, 3);
                button_wm_nl.Size = new Size(24, 24);
            }
            else
            {
                button_wm_nl.Location = new System.Drawing.Point(697, 5);
                button_wm_nl.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_refresh_ws1.Location = new System.Drawing.Point(3, 3);
                button_refresh_ws1.Size = new Size(24, 24);
            }
            else
            {
                button_refresh_ws1.Location = new System.Drawing.Point(5, 5);
                button_refresh_ws1.Size = new Size(21, 21);
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_load_weld_map.Location = new System.Drawing.Point(494, 2);
                button_load_weld_map.Size = new Size(198, 30);

                button_load_weld_map.Font = englargedFont;
            }
            else
            {
                button_load_weld_map.Location = new System.Drawing.Point(502, 2);
                button_load_weld_map.Size = new Size(189, 28);

                button_load_weld_map.Font = regularfont;
            }


            if (radioButton_enlarged.Checked == true)
            {
                comboBox_ws1.Location = new System.Drawing.Point(32, 4);
                comboBox_ws1.Size = new Size(455, 25);

                comboBox_ws1.Font = englargedFont;
            }
            else
            {
                comboBox_ws1.Location = new System.Drawing.Point(32, 4);
                comboBox_ws1.Size = new Size(464, 24);

                comboBox_ws1.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                panel6.Location = new System.Drawing.Point(3, 108);
                panel6.Size = new Size(723, 28);
            }
            else
            {
                panel6.Location = new System.Drawing.Point(3, 101);
                panel6.Size = new Size(723, 25);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label18.Location = new System.Drawing.Point(5, 5);
                label18.Size = new Size(93, 23);

                label18.Font = englargedHeader;
            }
            else
            {
                label18.Location = new System.Drawing.Point(3, 3);
                label18.Size = new Size(75, 18);

                label18.Font = regularHeader;
            }

            if (radioButton_enlarged.Checked == true)
            {
                dataGridView_error_weld_map.Location = new System.Drawing.Point(3, 134);
                dataGridView_error_weld_map.Size = new Size(723, 400);

                dataGridView_error_weld_map.DefaultCellStyle.Font = englargedFont;

                dataGridView_error_weld_map.RowHeadersDefaultCellStyle.Font = englargedHeader;

            }
            else
            {
                dataGridView_error_weld_map.Location = new System.Drawing.Point(3, 126);
                dataGridView_error_weld_map.Size = new Size(723, 427);

                dataGridView_error_weld_map.DefaultCellStyle.Font = regularfont;
                dataGridView_error_weld_map.RowHeadersDefaultCellStyle.Font = regularHeader;

            }

            if (radioButton_enlarged.Checked == true)
            {
                panel_stats.Location = new System.Drawing.Point(3, 540);
                panel_stats.Size = new Size(722, 130);
            }
            else
            {
                panel_stats.Location = new System.Drawing.Point(3, 555);
                panel_stats.Size = new Size(722, 114);
            }

            if (radioButton_enlarged.Checked == true)
            {
                label19.Location = new System.Drawing.Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = englargedFont;
            }
            else
            {
                label19.Location = new System.Drawing.Point(5, 3);
                label19.Size = new Size(58, 14);

                label19.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_Items.Location = new System.Drawing.Point(3, 25);
                textBox_PM_Items.Size = new Size(300, 25);

                textBox_PM_Items.Font = englargedFont;
            }
            else
            {
                textBox_PM_Items.Location = new System.Drawing.Point(3, 21);
                textBox_PM_Items.Size = new Size(287, 20);

                textBox_PM_Items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_rows.Location = new System.Drawing.Point(315, 25);
                textBox_PM_no_rows.Size = new Size(40, 25);

                textBox_PM_no_rows.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_rows.Location = new System.Drawing.Point(311, 21);
                textBox_PM_no_rows.Size = new Size(37, 20);

                textBox_PM_no_rows.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_missing_OD.Location = new System.Drawing.Point(3, 51);
                textBox_PM_missing_OD.Size = new Size(300, 25);

                textBox_PM_missing_OD.Font = englargedFont;
            }
            else
            {
                textBox_PM_missing_OD.Location = new System.Drawing.Point(3, 42);
                textBox_PM_missing_OD.Size = new Size(287, 20);

                textBox_PM_missing_OD.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_PM_no_duplicates.Location = new System.Drawing.Point(315, 51);
                textBox_PM_no_duplicates.Size = new Size(40, 25);

                textBox_PM_no_duplicates.Font = englargedFont;
            }
            else
            {
                textBox_PM_no_duplicates.Location = new System.Drawing.Point(311, 42);
                textBox_PM_no_duplicates.Size = new Size(37, 20);

                textBox_PM_no_duplicates.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_WM_defl_issues.Location = new System.Drawing.Point(3, 77);
                textBox_WM_defl_issues.Size = new Size(300, 25);

                textBox_WM_defl_issues.Font = englargedFont;
            }
            else
            {
                textBox_WM_defl_issues.Location = new System.Drawing.Point(3, 63);
                textBox_WM_defl_issues.Size = new Size(287, 20);

                textBox_WM_defl_issues.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_WM_no_defl.Location = new System.Drawing.Point(315, 77);
                textBox_WM_no_defl.Size = new Size(40, 25);

                textBox_WM_no_defl.Font = englargedFont;
            }
            else
            {
                textBox_WM_no_defl.Location = new System.Drawing.Point(311, 63);
                textBox_WM_no_defl.Size = new Size(37, 20);

                textBox_WM_no_defl.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_WM_null_value_items.Location = new System.Drawing.Point(3, 103);
                textBox_WM_null_value_items.Size = new Size(300, 25);

                textBox_WM_null_value_items.Font = englargedFont;
            }
            else
            {
                textBox_WM_null_value_items.Location = new System.Drawing.Point(3, 84);
                textBox_WM_null_value_items.Size = new Size(287, 20);

                textBox_WM_null_value_items.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                textBox_WM_no_null.Location = new System.Drawing.Point(315, 103);
                textBox_WM_no_null.Size = new Size(40, 25);

                textBox_WM_no_null.Font = englargedFont;
            }
            else
            {
                textBox_WM_no_null.Location = new System.Drawing.Point(311, 84);
                textBox_WM_no_null.Size = new Size(37, 20);

                textBox_WM_no_null.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_pipe_rep.Location = new System.Drawing.Point(503, 3);
                button_pipe_rep.Size = new Size(214, 30);

                button_pipe_rep.Font = englargedFont;
            }
            else
            {
                button_pipe_rep.Location = new System.Drawing.Point(556, 3);
                button_pipe_rep.Size = new Size(161, 28);

                button_pipe_rep.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_export_errors_to_xl.Location = new System.Drawing.Point(503, 96);
                button_export_errors_to_xl.Size = new Size(214, 30);

                button_export_errors_to_xl.Font = englargedFont;
            }
            else
            {
                button_export_errors_to_xl.Location = new System.Drawing.Point(556, 80);
                button_export_errors_to_xl.Size = new Size(161, 28);

                button_export_errors_to_xl.Font = regularfont;
            }

            if (radioButton_enlarged.Checked == true)
            {
                button_dismiss_errors.Location = new System.Drawing.Point(408, -1);
                button_dismiss_errors.Size = new Size(124, 30);

                button_dismiss_errors.Font = englargedFont;
            }
            else
            {
                button_dismiss_errors.Location = new System.Drawing.Point(408, -1);
                button_dismiss_errors.Size = new Size(104, 25);

                button_dismiss_errors.Font = regularfont;
            }
        }

        private void display_errors(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                dt1.Columns.RemoveAt(6);
                dt1.Columns.RemoveAt(5);


                if (dt_dismissed_errors != null && dt_dismissed_errors.Rows.Count > 0)
                {

                    for (int i = dt_display.Rows.Count - 1; i >= 0; --i)
                    {

                        string val0i = "";
                        string val1i = "";
                        string val2i = "";
                        string val4i = "";

                        if (dt_display.Rows[i][0] != DBNull.Value)
                        {
                            val0i = Convert.ToString(dt_display.Rows[i][0]);
                        }

                        if (dt_display.Rows[i][1] != DBNull.Value)
                        {
                            val1i = Convert.ToString(dt_display.Rows[i][1]);
                        }

                        if (dt_display.Rows[i][2] != DBNull.Value)
                        {
                            val2i = Convert.ToString(dt_display.Rows[i][2]);
                        }

                        if (dt_display.Rows[i][4] != DBNull.Value)
                        {
                            val4i = Convert.ToString(dt_display.Rows[i][4]);
                        }


                        for (int j = 0; j < dt_dismissed_errors.Rows.Count; ++j)
                        {

                            string val0j = "";
                            string val1j = "";
                            string val2j = "";
                            string val3j = "";


                            if (dt_dismissed_errors.Rows[j][0] != DBNull.Value)
                            {
                                val0j = Convert.ToString(dt_dismissed_errors.Rows[j][0]);
                            }

                            if (dt_dismissed_errors.Rows[j][1] != DBNull.Value)
                            {
                                val1j = Convert.ToString(dt_dismissed_errors.Rows[j][1]);
                            }

                            if (dt_dismissed_errors.Rows[j][2] != DBNull.Value)
                            {
                                val2j = Convert.ToString(dt_dismissed_errors.Rows[j][2]);
                            }

                            if (dt_dismissed_errors.Rows[j][3] != DBNull.Value)
                            {
                                val3j = Convert.ToString(dt_dismissed_errors.Rows[j][3]);
                            }

                            if (val0i == val0j && val1i == val1j && val2i == val2j && val4i == val3j)
                            {
                                dt_display.Rows[i].Delete();

                                j = dt_dismissed_errors.Rows.Count;
                            }



                        }


                    }

                }


                dataGridView_error_weld_map.DataSource = dt1;
                dataGridView_error_weld_map.Columns[0].Width = 75;
                dataGridView_error_weld_map.Columns[1].Width = 100;
                dataGridView_error_weld_map.Columns[2].Width = 200;
                dataGridView_error_weld_map.Columns[3].Width = 75;
                dataGridView_error_weld_map.Columns[4].Width = 300;
                dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_weld_map.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_weld_map.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_weld_map.EnableHeadersVisualStyles = false;
            }
            else
            {
                dataGridView_error_weld_map.DataSource = null;
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
                    dt_dismissed_errors.Columns.Add("Feature Code", typeof(string));
                    dt_dismissed_errors.Columns.Add("Value", typeof(string));
                    dt_dismissed_errors.Columns.Add("Error", typeof(string));
                }

                List<int> lista1 = new List<int>();

                foreach (DataGridViewCell cell1 in dataGridView_error_weld_map.SelectedCells)
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
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][0] = dataGridView_error_weld_map.Rows[lista1[i]].Cells[0].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][1] = dataGridView_error_weld_map.Rows[lista1[i]].Cells[1].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][2] = dataGridView_error_weld_map.Rows[lista1[i]].Cells[2].Value;
                        dt_dismissed_errors.Rows[dt_dismissed_errors.Rows.Count - 1][3] = dataGridView_error_weld_map.Rows[lista1[i]].Cells[4].Value;
                    }

                    if (W3 == null)
                    {
                        Functions.Create_a_new_worksheet_from_excel_by_name(Wgen_main_form.tpage_allpts.filename, dismiss_errors_tab);

                    }
                    Functions.Transfer_datatable_to_existing_excel_spreadsheet_by_name(dt_dismissed_errors, Wgen_main_form.tpage_allpts.filename, dismiss_errors_tab, false);

                    dt_display = dt_errors.Copy();

                    display_errors(dt_display);

                }


            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void Add_station_to_all_points(Polyline cl_poly)
        {
            if (Wgen_main_form.dt_all_points != null || Wgen_main_form.dt_all_points.Rows.Count >= 0)
            {
                if (Wgen_main_form.dt_all_points.Columns.Contains(col_sta_lin) == true)
                {
                    Wgen_main_form.dt_all_points.Columns.Remove(col_sta_lin);
                }
                if (Wgen_main_form.dt_all_points.Columns.Contains(col_sta_ifc) == true)
                {
                    Wgen_main_form.dt_all_points.Columns.Remove(col_sta_ifc);
                }
                Wgen_main_form.dt_all_points.Columns.Add(col_sta_lin, typeof(double));
                Wgen_main_form.dt_all_points.Columns.Add(col_sta_ifc, typeof(double));


                for (int i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                {
                    if (
                        Wgen_main_form.dt_all_points.Rows[i][colpt2] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2])) == true &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt3] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3])) == true)
                    {
                        double y = Convert.ToDouble(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2]));
                        double x = Convert.ToDouble(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3]));
                        double sta1 = cl_poly.GetDistAtPoint(cl_poly.GetClosestPointTo(new Point3d(x, y, cl_poly.Elevation), Vector3d.ZAxis, false));

                        Wgen_main_form.dt_all_points.Rows[i][col_sta_lin] = sta1;
                        if (dt_st_eq != null && dt_st_eq.Rows.Count > 0)
                        {
                            double sta_eq = Station_equation_of(sta1);
                            Wgen_main_form.dt_all_points.Rows[i][col_sta_ifc] = sta_eq;
                        }
                    }
                }
            }
        }

        private int get_dt_column_number_from_letter(string cuvant)
        {
            int idx = -1;

            if (cuvant.Length == 3 && cuvant.Contains("{") == true && cuvant.Contains("}") == true)
            {
                string letter1 = cuvant.Replace("{", "").Replace("}", "");
                idx = Functions.get_excel_column_index(letter1);
                return idx - 1;
            }



            return -1;
        }


        private void button_create_weldmapR2_Click(object sender, EventArgs e)
        {

            string pathR2 = Wgen_main_form.WGEN_folder + Wgen_main_form.wmr2;
            if (System.IO.File.Exists(pathR2) == false)
            {
                MessageBox.Show("no file found\r\n" + pathR2 + "\r\nOperation aborted");
                return;
            }

            if (Wgen_main_form.dt_all_points == null || Wgen_main_form.dt_all_points.Rows.Count == 0 || Wgen_main_form.dt_ground_tally == null || Wgen_main_form.dt_ground_tally.Rows.Count == 0)
            {
                return;
            }

            set_enable_false();

            #region calc station


            set_enable_false();
            int i = 0;
            dt_st_eq = null;
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {

                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        if (InputBox("WGEN", "Select the Centerline:") == DialogResult.OK)
                        {

                        }
                        else
                        {
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            return;
                        }

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt5;
                        Prompt5 = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt5.SetRejectMessage("\nSelect a polyline!");
                        Prompt5.AllowNone = true;
                        Prompt5.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Prompt5.AddAllowedClass(typeof(Autodesk.Civil.DatabaseServices.Alignment), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt5);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            MessageBox.Show("no centerline");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            return;
                        }

                        Polyline Poly_cl = null;
                        Poly_cl = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        bool is_from_alignment = false;
                        if (Poly_cl == null)
                        {
                            Autodesk.Civil.DatabaseServices.Alignment align1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Autodesk.Civil.DatabaseServices.Alignment;
                            if (align1 != null)
                            {
                                Poly_cl = Trans1.GetObject(align1.GetPolyline(), OpenMode.ForRead) as Polyline;

                                if (Poly_cl == null)
                                {
                                    MessageBox.Show("error at the alignment object");
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    return;
                                }

                                double Start1 = align1.StartingStation;
                                is_from_alignment = true;
                                Autodesk.Civil.DatabaseServices.StationEquationCollection col_st_eq = align1.StationEquations;

                                if (col_st_eq.Count > 0 || Start1 != 0)
                                {
                                    dt_st_eq = new System.Data.DataTable();
                                    dt_st_eq.Columns.Add("Back", typeof(double));
                                    dt_st_eq.Columns.Add("Ahead", typeof(double));
                                    if (Start1 != 0)
                                    {
                                        dt_st_eq.Rows.Add();
                                        dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Back"] = 0;
                                        dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Ahead"] = Start1;
                                    }
                                    if (col_st_eq.Count > 0)
                                    {
                                        for (i = 0; i < col_st_eq.Count; ++i)
                                        {
                                            Autodesk.Civil.DatabaseServices.StationEquation eq1 = col_st_eq[i];
                                            dt_st_eq.Rows.Add();
                                            dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Back"] = eq1.StationBack;
                                            dt_st_eq.Rows[dt_st_eq.Rows.Count - 1]["Ahead"] = eq1.StationAhead;
                                        }
                                    }
                                    dt_st_eq = Functions.Sort_data_table(dt_st_eq, "Back");
                                }
                            }
                        }

                        if (Poly_cl != null)
                        {
                            Add_station_to_all_points(Poly_cl);
                            if (is_from_alignment == true)
                            {
                                BTrecord.UpgradeOpen();
                                Poly_cl.UpgradeOpen();
                                Poly_cl.Erase();
                            }
                        }
                        else
                        {
                            MessageBox.Show("no centerline");
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }


                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                set_enable_true();
                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                this.MdiParent.WindowState = FormWindowState.Normal;
                set_enable_true();

                return;
            }

            this.MdiParent.WindowState = FormWindowState.Normal;

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            #endregion

            System.Data.DataTable dt_title = null;
            System.Data.DataTable dt_legend = null;
            System.Data.DataTable dt_fc_ap = null;
            System.Data.DataTable dt_fc_gt = null;

            load_R2_tables(pathR2, ref dt_title, ref dt_legend, ref dt_fc_ap, ref dt_fc_gt);

            System.Data.DataTable dt_wmR2 = new System.Data.DataTable();
            dt_wmR2.Columns.Add("delete", typeof(string));

            dt_wmR2.Columns.Add(col_station, typeof(double));
            dt_wmR2.Columns.Add(col_northing, typeof(double));
            dt_wmR2.Columns.Add(col_easting, typeof(double));
            dt_wmR2.Columns.Add(col_elevation, typeof(double));
            dt_wmR2.Columns.Add(col_description, typeof(string));
            dt_wmR2.Columns.Add(col_type, typeof(string));
            dt_wmR2.Columns.Add(col_7, typeof(string));
            dt_wmR2.Columns.Add(col_8, typeof(string));
            dt_wmR2.Columns.Add(col_9, typeof(string));
            dt_wmR2.Columns.Add(col_10, typeof(string));
            dt_wmR2.Columns.Add(col_11, typeof(string));
            dt_wmR2.Columns.Add(col_12, typeof(string));
            dt_wmR2.Columns.Add(col_13, typeof(string));
            dt_wmR2.Columns.Add(col_14, typeof(string));
            dt_wmR2.Columns.Add(col_15, typeof(string));
            dt_wmR2.Columns.Add(col_16, typeof(string));
            dt_wmR2.Columns.Add(col_17, typeof(string));

            System.Data.DataTable dt_ap = Wgen_main_form.dt_all_points;
            System.Data.DataTable dt_gt = Wgen_main_form.dt_ground_tally;

            try
            {
                dt_ap = Functions.Sort_data_table(Wgen_main_form.dt_all_points, colpt_sta);


                // Functions.Transfer_datatable_to_new_excel_spreadsheet(dt_ap, "all points");



                Wgen_main_form.tpage_weldmap.Hide();
                Wgen_main_form.tpage_blank.Show();
                Wgen_main_form.tpage_pipe_manifest.Hide();
                Wgen_main_form.tpage_pipe_tally.Hide();
                Wgen_main_form.tpage_allpts.Hide();
                Wgen_main_form.tpage_build_pipe_tally.Hide();
                Wgen_main_form.tpage_duplicates.Hide();
                Wgen_main_form.tpage_blank.get_label_wait_visible(true);

                this.Refresh();


                bool is_station_eq = false;

                if (dt_st_eq != null && dt_st_eq.Rows.Count > 0)
                {
                    is_station_eq = true;
                }


                string col_mm_back = colpt9;
                string col_mm_ahead = colpt10;

                dt_ap.TableName = "ALLPTS";
                dt_gt.TableName = "PIPETALLY";


                DataSet dataset1 = new DataSet();
                dataset1.Tables.Add(dt_ap);
                dataset1.Tables.Add(dt_gt);

                DataRelation relation_mmid_back = new DataRelation("xxx", dt_ap.Columns[col_mm_back], dt_gt.Columns[colgt1], false);
                dataset1.Relations.Add(relation_mmid_back);

                DataRelation relation_mmid_ahead = new DataRelation("xxx1", dt_ap.Columns[col_mm_ahead], dt_gt.Columns[colgt1], false);
                dataset1.Relations.Add(relation_mmid_ahead);


                for (i = 0; i < dt_ap.Rows.Count; ++i)
                {
                    if (dt_ap.Rows[i][colpt5] != DBNull.Value)
                    {
                        string fc1 = Convert.ToString(dt_ap.Rows[i][colpt5]);

                        for (int k = 0; k < dt_fc_ap.Rows.Count; ++k)
                        {
                            if (dt_fc_ap.Rows[k][col_description] != DBNull.Value)
                            {
                                string fc2 = Convert.ToString(dt_fc_ap.Rows[k][col_description]);
                                if (fc1.ToUpper() == fc2.ToUpper())
                                {
                                    bool add_row = false;

                                    for (int j = 1; j < dt_fc_ap.Columns.Count; ++j)
                                    {
                                        if (dt_fc_ap.Rows[k][j] != DBNull.Value)
                                        {
                                            add_row = true;
                                            j = dt_fc_ap.Columns.Count;
                                        }
                                    }

                                    if (add_row == true)
                                    {
                                        dt_wmR2.Rows.Add();
                                        if (is_station_eq == false)
                                        {
                                            dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_station] = dt_ap.Rows[i][col_sta_lin];
                                        }
                                        else
                                        {
                                            dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_station] = dt_ap.Rows[i][col_sta_ifc];
                                        }


                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1]["delete"] = dt_ap.Rows[i][0];
                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_northing] = dt_ap.Rows[i][colpt2];
                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_easting] = dt_ap.Rows[i][colpt3];
                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_elevation] = dt_ap.Rows[i][colpt4];
                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_description] = dt_ap.Rows[i][colpt5];



                                        int col_ap_start = dt_fc_ap.Columns.IndexOf(col_type);
                                        int col_ap_end = dt_fc_ap.Columns.Count - 1;
                                        int col_index_r2 = dt_wmR2.Columns.IndexOf(col_type);

                                        for (int m = col_ap_start; m <= col_ap_end; ++m)
                                        {
                                            if (dt_fc_ap.Rows[k][m] != DBNull.Value)
                                            {
                                                string val1 = Convert.ToString(dt_fc_ap.Rows[k][m]);
                                                string cell_value = "";
                                                if (val1.Contains("{") == true && val1.Contains("}") == true)
                                                {
                                                    do
                                                    {
                                                        int index1 = val1.IndexOf("{", 0);
                                                        int index2 = val1.IndexOf("}", 0);
                                                        string column1 = val1.Substring(index1, index2 - index1 + 1);
                                                        int col_idx = get_dt_column_number_from_letter(column1);
                                                        string content1 = "";
                                                        if (dt_ap.Rows[i][col_idx] != DBNull.Value)
                                                        {
                                                            content1 = Convert.ToString(dt_ap.Rows[i][col_idx]);
                                                        }
                                                        val1 = val1.Replace(column1, content1);
                                                        cell_value = val1;
                                                    } while (val1.Contains("{") == true);
                                                }

                                                if (cell_value != "")
                                                {
                                                    dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_index_r2] = cell_value;
                                                }
                                            }
                                            else
                                            {
                                                if (dt_fc_gt.Rows[k][m] != DBNull.Value)
                                                {
                                                    string val1 = Convert.ToString(dt_fc_gt.Rows[k][m]);
                                                    string cell_value = "";
                                                    if (val1.Contains("{") == true && val1.Contains("}") == true)
                                                    {
                                                        do
                                                        {
                                                            int index1 = val1.IndexOf("{", 0);
                                                            int index2 = val1.IndexOf("}", 0);
                                                            string column1 = val1.Substring(index1, index2 - index1 + 1);
                                                            bool use_back = false;
                                                            bool use_all_points = false;

                                                            string content1 = "";
                                                            if (column1.Contains((char)34 + "all_points_column_") == true && column1.Contains("_with_") == true)
                                                            {
                                                                use_back = true;

                                                                string string1 = column1.Replace((char)34 + "all_points_column_", "").Replace("{", "").Replace("}", "");
                                                                int pos1 = string1.IndexOf("_");
                                                                string col_ap = string1.Substring(0, pos1);
                                                                string col_gt = string1.Replace(col_ap + "_with_", "");

                                                                col_ap = "{" + col_ap + "}";
                                                                col_gt = "{" + col_gt + "}";

                                                                int col_ap_index = get_dt_column_number_from_letter(col_ap);
                                                                int col_gt_index = get_dt_column_number_from_letter(col_gt);

                                                                if (dt_ap.Rows[i][col_ap_index] != DBNull.Value)
                                                                {
                                                                    string val_from_ap = Convert.ToString(dt_ap.Rows[i][col_ap_index]);

                                                                    for (int q = 0; q < dt_gt.Rows.Count; ++q)
                                                                    {
                                                                        string val_from_gt = Convert.ToString(dt_gt.Rows[q][colgt1]);
                                                                        if (val_from_ap == val_from_gt)
                                                                        {
                                                                            content1 = Convert.ToString(dt_gt.Rows[q][col_gt_index]);
                                                                            q = dt_gt.Rows.Count;
                                                                        }

                                                                    }

                                                                }



                                                                val1 = "";
                                                            }
                                                            if (column1.Contains((char)34 + "all_points" + (char)34) == true)
                                                            {
                                                                use_all_points = true;
                                                                column1 = column1.Replace((char)34 + "all_points" + (char)34, "");
                                                                val1 = val1.Replace((char)34 + "all_points" + (char)34, "");
                                                            }

                                                            int col_idx = get_dt_column_number_from_letter(column1);


                                                            if (use_all_points == false)
                                                            {
                                                                if (use_back == false && dt_ap.Rows[i].GetChildRows(relation_mmid_ahead).Length > 0)
                                                                {
                                                                    System.Data.DataRow row1 = dt_ap.Rows[i].GetChildRows(relation_mmid_ahead)[0];
                                                                    if (row1[col_idx] != DBNull.Value)
                                                                    {
                                                                        content1 = Convert.ToString(row1[col_idx]);
                                                                    }
                                                                }

                                                            }


                                                            if (use_back == false && use_all_points == true)
                                                            {
                                                                if (dt_ap.Rows[i][col_idx] != DBNull.Value)
                                                                {
                                                                    content1 = Convert.ToString(dt_ap.Rows[i][col_idx]);
                                                                }
                                                            }

                                                            if (val1 != "")
                                                            {
                                                                val1 = val1.Replace(column1, content1);
                                                            }
                                                            else
                                                            {
                                                                val1 = content1;
                                                            }

                                                            cell_value = val1;
                                                        } while (val1.Contains("{") == true);
                                                    }

                                                    if (cell_value != "")
                                                    {
                                                        dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][col_index_r2] = cell_value;
                                                    }
                                                }
                                            }
                                            ++col_index_r2;
                                        }



                                        if (dt_wmR2.Columns.Contains(legend_cover_column) == true)
                                        {
                                            dt_wmR2.Rows[dt_wmR2.Rows.Count - 1][legend_cover_column] = dt_ap.Rows[i]["Cover"];

                                        }

                                    }
                                }
                            }
                        }
                    }

                }

                transfer_to_excel_wm_R2(dt_title, dt_legend, dt_wmR2);
               // Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_ap);


                dataset1.Relations.Remove(relation_mmid_back);
                dataset1.Relations.Remove(relation_mmid_ahead);
                dataset1.Tables.Remove(dt_ap);
                dataset1.Tables.Remove(dt_gt);




            }
            catch (System.Exception ex)
            {
                MessageBox.Show(i.ToString() + ":\r\n" + ex.Message);
            }
            set_enable_true();

            Wgen_main_form.tpage_weldmap.Show();
            Wgen_main_form.tpage_blank.Hide();
            Wgen_main_form.tpage_pipe_manifest.Hide();
            Wgen_main_form.tpage_pipe_tally.Hide();
            Wgen_main_form.tpage_allpts.Hide();
            Wgen_main_form.tpage_build_pipe_tally.Hide();
            Wgen_main_form.tpage_duplicates.Hide();
            Wgen_main_form.tpage_blank.get_label_wait_visible(false);

        }





        private void load_R2_tables(string file1, ref System.Data.DataTable dt0, ref System.Data.DataTable dt1, ref System.Data.DataTable dt2, ref System.Data.DataTable dt3)
        {
            dt0 = new System.Data.DataTable();
            dt0.Columns.Add();

            dt1 = new System.Data.DataTable();
            dt1.Columns.Add(col_description, typeof(string));
            dt1.Columns.Add(col_type, typeof(string));
            dt1.Columns.Add(col_7, typeof(string));
            dt1.Columns.Add(col_8, typeof(string));
            dt1.Columns.Add(col_9, typeof(string));
            dt1.Columns.Add(col_10, typeof(string));
            dt1.Columns.Add(col_11, typeof(string));
            dt1.Columns.Add(col_12, typeof(string));
            dt1.Columns.Add(col_13, typeof(string));
            dt1.Columns.Add(col_14, typeof(string));
            dt1.Columns.Add(col_15, typeof(string));
            dt1.Columns.Add(col_16, typeof(string));
            dt1.Columns.Add(col_17, typeof(string));


            dt2 = new System.Data.DataTable();
            dt2 = dt1.Clone();

            dt3 = new System.Data.DataTable();
            dt3 = dt1.Clone();

            bool excel_is_opened = false;
            bool file_is_opened = false;

            if (System.IO.File.Exists(file1) == true)
            {
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = null;
                try
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workbook2.Worksheets)
                            {
                                if (Workbook2.FullName == file1)
                                {
                                    file_is_opened = true;

                                    Workbook1 = Workbook2;

                                }
                            }
                        }
                        Excel1.Visible = true;
                        excel_is_opened = true;
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                        Excel1.Visible = false;
                    }

                    if (file_is_opened == false)
                    {
                        Workbook1 = Excel1.Workbooks.Open(file1);
                    }

                    foreach (Microsoft.Office.Interop.Excel.Worksheet W1 in Workbook1.Worksheets)
                    {
                        string nume1 = W1.Name;
                        if (nume1.ToUpper() == "FC_LEGEND")
                        {
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A2: M30001"];
                            object[,] values1 = new object[30000, 13];
                            values1 = range1.Value2;
                            for (int i = 1; i <= 30000; ++i)
                            {
                                object valA = values1[i, 1];
                                object valB = values1[i, 2];
                                object valC = values1[i, 3];
                                object valD = values1[i, 4];
                                object valE = values1[i, 5];
                                object valF = values1[i, 6];
                                object valG = values1[i, 7];
                                object valH = values1[i, 8];
                                object valI = values1[i, 9];
                                object valJ = values1[i, 10];
                                object valK = values1[i, 11];
                                object valL = values1[i, 12];
                                object valM = values1[i, 13];




                                if (valA != null)
                                {

                                    string fc = Convert.ToString(valA);
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][col_description] = fc;
                                    if (valB != null) dt1.Rows[dt1.Rows.Count - 1][col_type] = Convert.ToString(valB);
                                    if (valC != null) dt1.Rows[dt1.Rows.Count - 1][col_7] = Convert.ToString(valC);
                                    if (valD != null) dt1.Rows[dt1.Rows.Count - 1][col_8] = Convert.ToString(valD);
                                    if (valE != null) dt1.Rows[dt1.Rows.Count - 1][col_9] = Convert.ToString(valE);
                                    if (valF != null) dt1.Rows[dt1.Rows.Count - 1][col_10] = Convert.ToString(valF);
                                    if (valG != null) dt1.Rows[dt1.Rows.Count - 1][col_11] = Convert.ToString(valG);
                                    if (valH != null) dt1.Rows[dt1.Rows.Count - 1][col_12] = Convert.ToString(valH);
                                    if (valI != null) dt1.Rows[dt1.Rows.Count - 1][col_13] = Convert.ToString(valI);
                                    if (valJ != null) dt1.Rows[dt1.Rows.Count - 1][col_14] = Convert.ToString(valJ);
                                    if (valK != null) dt1.Rows[dt1.Rows.Count - 1][col_15] = Convert.ToString(valK);
                                    if (valL != null) dt1.Rows[dt1.Rows.Count - 1][col_16] = Convert.ToString(valL);
                                    if (valM != null) dt1.Rows[dt1.Rows.Count - 1][col_17] = Convert.ToString(valM);

                                    if (fc.ToUpper() == "WELD")
                                    {
                                        if (valB != null) if (Convert.ToString(valB).ToUpper() == "COVER") legend_cover_column = col_type;
                                        if (valC != null) if (Convert.ToString(valC).ToUpper() == "COVER") legend_cover_column = col_7;
                                        if (valD != null) if (Convert.ToString(valD).ToUpper() == "COVER") legend_cover_column = col_8;
                                        if (valE != null) if (Convert.ToString(valE).ToUpper() == "COVER") legend_cover_column = col_9;
                                        if (valF != null) if (Convert.ToString(valF).ToUpper() == "COVER") legend_cover_column = col_10;
                                        if (valG != null) if (Convert.ToString(valG).ToUpper() == "COVER") legend_cover_column = col_11;
                                        if (valH != null) if (Convert.ToString(valH).ToUpper() == "COVER") legend_cover_column = col_12;
                                        if (valI != null) if (Convert.ToString(valI).ToUpper() == "COVER") legend_cover_column = col_13;
                                        if (valJ != null) if (Convert.ToString(valJ).ToUpper() == "COVER") legend_cover_column = col_14;
                                        if (valK != null) if (Convert.ToString(valK).ToUpper() == "COVER") legend_cover_column = col_15;
                                        if (valL != null) if (Convert.ToString(valL).ToUpper() == "COVER") legend_cover_column = col_16;
                                        if (valM != null) if (Convert.ToString(valM).ToUpper() == "COVER") legend_cover_column = col_17;
                                    }


                                }
                                else
                                {
                                    i = values1.Length + 1;
                                }
                            }


                        }

                        if (nume1.ToUpper() == "AP_COLUMNS")
                        {
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A2: M30001"];
                            object[,] values1 = new object[30000, 13];
                            values1 = range1.Value2;
                            for (int i = 1; i <= 30000; ++i)
                            {
                                object valA = values1[i, 1];
                                object valB = values1[i, 2];
                                object valC = values1[i, 3];
                                object valD = values1[i, 4];
                                object valE = values1[i, 5];
                                object valF = values1[i, 6];
                                object valG = values1[i, 7];
                                object valH = values1[i, 8];
                                object valI = values1[i, 9];
                                object valJ = values1[i, 10];
                                object valK = values1[i, 11];
                                object valL = values1[i, 12];
                                object valM = values1[i, 13];




                                if (valA != null)
                                {
                                    string fc = Convert.ToString(valA);
                                    dt2.Rows.Add();
                                    dt2.Rows[dt2.Rows.Count - 1][col_description] = fc;
                                    if (valB != null) dt2.Rows[dt2.Rows.Count - 1][col_type] = Convert.ToString(valB);
                                    if (valC != null) dt2.Rows[dt2.Rows.Count - 1][col_7] = Convert.ToString(valC);
                                    if (valD != null) dt2.Rows[dt2.Rows.Count - 1][col_8] = Convert.ToString(valD);
                                    if (valE != null) dt2.Rows[dt2.Rows.Count - 1][col_9] = Convert.ToString(valE);
                                    if (valF != null) dt2.Rows[dt2.Rows.Count - 1][col_10] = Convert.ToString(valF);
                                    if (valG != null) dt2.Rows[dt2.Rows.Count - 1][col_11] = Convert.ToString(valG);
                                    if (valH != null) dt2.Rows[dt2.Rows.Count - 1][col_12] = Convert.ToString(valH);
                                    if (valI != null) dt2.Rows[dt2.Rows.Count - 1][col_13] = Convert.ToString(valI);
                                    if (valJ != null) dt2.Rows[dt2.Rows.Count - 1][col_14] = Convert.ToString(valJ);
                                    if (valK != null) dt2.Rows[dt2.Rows.Count - 1][col_15] = Convert.ToString(valK);
                                    if (valL != null) dt2.Rows[dt2.Rows.Count - 1][col_16] = Convert.ToString(valL);
                                    if (valM != null) dt2.Rows[dt2.Rows.Count - 1][col_17] = Convert.ToString(valM);

                                    //if (fc.ToUpper() == "WELD" && ap_cover_column_XL == "XXX")
                                    //{
                                    //    if (valB != null) if (legend_cover_column == col_type) ap_cover_column_XL = Convert.ToString(valB).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valC != null) if (legend_cover_column == col_7) ap_cover_column_XL = Convert.ToString(valC).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valD != null) if (legend_cover_column == col_8) ap_cover_column_XL = Convert.ToString(valD).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valE != null) if (legend_cover_column == col_9) ap_cover_column_XL = Convert.ToString(valE).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valF != null) if (legend_cover_column == col_10) ap_cover_column_XL = Convert.ToString(valF).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valG != null) if (legend_cover_column == col_11) ap_cover_column_XL = Convert.ToString(valG).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valH != null) if (legend_cover_column == col_12) ap_cover_column_XL = Convert.ToString(valH).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valI != null) if (legend_cover_column == col_13) ap_cover_column_XL = Convert.ToString(valI).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valJ != null) if (legend_cover_column == col_14) ap_cover_column_XL = Convert.ToString(valJ).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valK != null) if (legend_cover_column == col_15) ap_cover_column_XL = Convert.ToString(valK).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valL != null) if (legend_cover_column == col_16) ap_cover_column_XL = Convert.ToString(valL).Replace("{", "").Replace("}", "").ToUpper();
                                    //    if (valM != null) if (legend_cover_column == col_17) ap_cover_column_XL = Convert.ToString(valM).Replace("{", "").Replace("}", "").ToUpper();
                                    //}

                                }
                                else
                                {
                                    i = values1.Length + 1;
                                }
                            }


                        }

                        if (nume1.ToUpper() == "GT_COLUMNS")
                        {
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A2: M30001"];
                            object[,] values1 = new object[30000, 13];
                            values1 = range1.Value2;
                            for (int i = 1; i <= 30000; ++i)
                            {
                                object valA = values1[i, 1];
                                object valB = values1[i, 2];
                                object valC = values1[i, 3];
                                object valD = values1[i, 4];
                                object valE = values1[i, 5];
                                object valF = values1[i, 6];
                                object valG = values1[i, 7];
                                object valH = values1[i, 8];
                                object valI = values1[i, 9];
                                object valJ = values1[i, 10];
                                object valK = values1[i, 11];
                                object valL = values1[i, 12];
                                object valM = values1[i, 13];




                                if (valA != null)
                                {
                                    dt3.Rows.Add();
                                    dt3.Rows[dt3.Rows.Count - 1][col_description] = Convert.ToString(valA);
                                    if (valB != null) dt3.Rows[dt3.Rows.Count - 1][col_type] = Convert.ToString(valB);
                                    if (valC != null) dt3.Rows[dt3.Rows.Count - 1][col_7] = Convert.ToString(valC);
                                    if (valD != null) dt3.Rows[dt3.Rows.Count - 1][col_8] = Convert.ToString(valD);
                                    if (valE != null) dt3.Rows[dt3.Rows.Count - 1][col_9] = Convert.ToString(valE);
                                    if (valF != null) dt3.Rows[dt3.Rows.Count - 1][col_10] = Convert.ToString(valF);
                                    if (valG != null) dt3.Rows[dt3.Rows.Count - 1][col_11] = Convert.ToString(valG);
                                    if (valH != null) dt3.Rows[dt3.Rows.Count - 1][col_12] = Convert.ToString(valH);
                                    if (valI != null) dt3.Rows[dt3.Rows.Count - 1][col_13] = Convert.ToString(valI);
                                    if (valJ != null) dt3.Rows[dt3.Rows.Count - 1][col_14] = Convert.ToString(valJ);
                                    if (valK != null) dt3.Rows[dt3.Rows.Count - 1][col_15] = Convert.ToString(valK);
                                    if (valL != null) dt3.Rows[dt3.Rows.Count - 1][col_16] = Convert.ToString(valL);
                                    if (valM != null) dt3.Rows[dt3.Rows.Count - 1][col_17] = Convert.ToString(valM);


                                }
                                else
                                {
                                    i = values1.Length + 1;
                                }
                            }


                        }

                        if (nume1.ToUpper() == "TITLE_INFO")
                        {
                            Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:A100"];
                            object[,] values1 = new object[100, 1];
                            values1 = range1.Value2;
                            for (int i = 1; i <= 101; ++i)
                            {
                                object valA = values1[i, 1];


                                if (valA != null)
                                {
                                    dt0.Rows.Add();
                                    dt0.Rows[dt0.Rows.Count - 1][0] = Convert.ToString(valA);


                                }
                                else
                                {
                                    i = values1.Length + 1;
                                }
                            }


                        }
                    }


                    if (file_is_opened == false) Workbook1.Close();
                    if (excel_is_opened == false) Excel1.Quit();

                }
                catch (System.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (excel_is_opened == false && Excel1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }




        }


        private void transfer_to_excel_wm_R2(System.Data.DataTable dt0, System.Data.DataTable dt1, System.Data.DataTable dt2)
        {
            if (dt0 != null && dt1 != null && dt2 != null)
            {
                if (dt1.Rows.Count > 0 && dt2.Rows.Count > 0)
                {

                    for (int i = dt1.Rows.Count - 1; i >= 0; --i)
                    {
                        bool delete = true;

                        for (int j = 1; j < dt1.Columns.Count; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                delete = false;
                                j = dt1.Columns.Count;
                            }
                        }

                        if (delete == false)
                        {
                            bool is_found = false;
                            string fc1 = Convert.ToString(dt1.Rows[i][col_description]);
                            for (int k = 0; k < dt2.Rows.Count; ++k)
                            {
                                if (dt2.Rows[k][col_description] != DBNull.Value)
                                {
                                    string fc2 = Convert.ToString(dt2.Rows[k][col_description]);
                                    if (fc1.ToUpper() == fc2.ToUpper())
                                    {
                                        is_found = true;
                                        k = dt2.Rows.Count;
                                    }
                                }
                            }
                            if (is_found == false) delete = true;
                        }


                        if (delete == true)
                        {
                            dt1.Rows[i].Delete();
                        }
                    }


                    Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_NEW_worksheet_from_Excel();
                    W1.Range["A:Q"].NumberFormat = "General";
                    W1.Range["A:A"].ColumnWidth = 12;
                    W1.Range["B:D"].ColumnWidth = 15;
                    W1.Range["E:E"].ColumnWidth = 23;
                    W1.Range["F:L"].ColumnWidth = 20;
                    W1.Range["M:M"].ColumnWidth = 75;
                    W1.Range["N:Q"].ColumnWidth = 18;


                    int start0 = 1;

                    for (int i = 0; i < dt0.Rows.Count; ++i)
                    {

                        Range rangeT = W1.Range["A" + Convert.ToString(i + 1)];
                        string value0 = Convert.ToString(dt0.Rows[i][0]);
                        if (value0.ToUpper().Contains("{DATE}") == true)
                        {
                            value0 = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
                        }

                        rangeT.Value2 = value0;

                        if (i < 2)
                        {
                            rangeT.Font.Size = 16;
                        }
                        else
                        {
                            rangeT.Font.Size = 12;
                        }


                        rangeT.Font.Name = "Arial";
                        if (i < 2) rangeT.Font.Bold = true;

                        rangeT.WrapText = false;
                        rangeT.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
                        rangeT.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                        rangeT.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
                        rangeT.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
                        rangeT.Interior.Pattern = 1;

                        W1.Range["A" + Convert.ToString(i + 1) + ":Q" + Convert.ToString(i + 1)].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        W1.Range["A" + Convert.ToString(i + 1) + ":Q" + Convert.ToString(i + 1)].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        W1.Range["A" + Convert.ToString(i + 1) + ":Q" + Convert.ToString(i + 1)].MergeCells = true;
                    }


                    int start1 = start0 + dt0.Rows.Count + 1;

                    int maxRows1 = dt1.Rows.Count;
                    int maxCols1 = dt1.Columns.Count;
                    string lastcol1 = Functions.get_excel_column_letter(maxCols1 + 4);
                    Range range1 = W1.Range["E" + Convert.ToString(start1 + 1) + ":" + lastcol1 + Convert.ToString(maxRows1 + start1)];
                    object[,] values1 = new object[maxRows1, maxCols1];

                    for (int i = 0; i < maxRows1; ++i)
                    {
                        for (int j = 0; j < maxCols1; ++j)
                        {
                            if (dt1.Rows[i][j] != DBNull.Value)
                            {
                                values1[i, j] = Convert.ToString(dt1.Rows[i][j]);
                            }
                        }
                    }

                    for (int j = 0; j < maxCols1; ++j)
                    {
                        string current1 = Functions.get_excel_column_letter(j + 5);

                        Range rangeT = W1.Range[current1 + Convert.ToString(start1)];
                        rangeT.Value2 = dt1.Columns[j].ColumnName;
                        rangeT.Interior.Color = 5645834;
                        rangeT.Font.Color = 16777215;
                        rangeT.Font.Size = 12;
                        rangeT.Font.Name = "Arial";
                        rangeT.Font.Bold = true;
                        rangeT.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        rangeT.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    }

                    range1.Value2 = values1;
                    range1.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    range1.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    range1.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    range1.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    range1.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
                    range1.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                    range1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    for (int i = 0; i < maxRows1; ++i)
                    {
                        if (dt1.Rows[i][0] != DBNull.Value)
                        {
                            string fc = Convert.ToString(dt1.Rows[i][0]);
                            if (fc.ToUpper() == "VALVE")
                            {
                                Range rangeX = W1.Range["E" + Convert.ToString(start1 + 1 + i) + ":" + lastcol1 + Convert.ToString(start1 + 1 + i)];
                                rangeX.Interior.Color = 15057582;
                            }
                        }
                    }

                    int start2 = start1 + dt1.Rows.Count + 2;

                    if (dt2.Columns.Contains("delete") == true)
                    {
                        dt2.Columns.Remove("delete");
                    }

                    int maxRows2 = dt2.Rows.Count;
                    int maxCols2 = dt2.Columns.Count;
                    string lastcol2 = Functions.get_excel_column_letter(maxCols2);
                    Range range2 = W1.Range["A" + Convert.ToString(start2 + 1) + ":" + lastcol2 + Convert.ToString(maxRows2 + start2)];
                    object[,] values2 = new object[maxRows2, maxCols2];

                    for (int i = 0; i < maxRows2; ++i)
                    {
                        for (int j = 0; j < maxCols2; ++j)
                        {
                            if (dt2.Rows[i][j] != DBNull.Value)
                            {
                                values2[i, j] = Convert.ToString(dt2.Rows[i][j]);
                            }
                        }
                    }

                    for (int j = 0; j < maxCols2; ++j)
                    {
                        string current1 = Functions.get_excel_column_letter(j + 1);

                        Range rangeT = W1.Range[current1 + Convert.ToString(start2)];
                        rangeT.Value2 = dt2.Columns[j].ColumnName;
                        rangeT.Interior.Color = 5645834;
                        rangeT.Font.Color = 16777215;
                        rangeT.Font.Size = 12;
                        rangeT.Font.Name = "Arial";
                        rangeT.Font.Bold = true;
                        rangeT.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        rangeT.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    }
                    range2.Value2 = values2;


                    range2.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    range2.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    range2.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    range2.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    range2.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
                    range2.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
                    range2.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    range2.VerticalAlignment = XlVAlign.xlVAlignCenter;

                    W1.Range["A" + Convert.ToString(start2 + 1) + ":A" + Convert.ToString(maxRows2 + start2)].NumberFormat = "0+00";
                    W1.Range["B" + Convert.ToString(start2 + 1) + ":D" + Convert.ToString(maxRows2 + start2)].NumberFormat = "0.000";

                    for (int i = 0; i < maxRows2; ++i)
                    {
                        if (dt2.Rows[i][col_description] != DBNull.Value)
                        {
                            string fc = Convert.ToString(dt2.Rows[i][col_description]);
                            if (fc.ToUpper() == "VALVE")
                            {
                                Range rangeX = W1.Range["A" + Convert.ToString(start2 + 1 + i) + ":" + lastcol2 + Convert.ToString(start2 + 1 + i)];
                                rangeX.Interior.Color = 15057582;
                            }
                        }
                    }
                    W1.Name = "Weld Map";

                }
            }

        }
        public void set_button_gen_wmR2()
        {
            if (Wgen_main_form.client_name.ToUpper() == "WHITEWATER")
            {
                button_gen_wmR2.Visible = true;
            }
            else
            {
                button_gen_wmR2.Visible = false;
            }

        }

    }
}
