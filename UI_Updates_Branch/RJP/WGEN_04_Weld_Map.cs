using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Alignment_mdi
{
    public partial class Wgen_weldmap : Form
    {
        private ContextMenuStrip ContextMenuStrip_go_to_error;

        System.Data.DataTable dt_errors;
        System.Data.DataTable dt_st_eq;

        double ng_tolerance = 100;
        Microsoft.Office.Interop.Excel.Worksheet W2 = null;

        int start_row = 2;
        string col1 = "PNT";
        string col2 = "NORTHING";
        string col3 = "EASTING";
        string col4 = "ELEVATION";
        string col5 = "FEATURE_CODE";
        string col6 = "DESCRIPTION";
        string col7 = "STATION_LINEAR";
        string col8 = "STATION_IFC";
        string col9 = "MM_BK";
        string col10 = "WALL_BK";
        string col11 = "PIPE_BK";
        string col12 = "HEAT_BK";
        string col13 = "COATING_BK";
        string col14 = "GRADE_BK";
        string col15 = "MM_AHD";
        string col16 = "WALL_AHD";
        string col17 = "PIPE_AHD";
        string col18 = "HEAT_AHD";
        string col19 = "COATING_AHD";
        string col20 = "GRADE_AHD";
        string col21 = "NG";
        string col22 = "NG_NORTHING";
        string col23 = "NG_EASTING";
        string col24 = "NG_ELEVATION";
        string col25 = "COVER";
        string col26 = "LOCATION";
        string col27 = "FILENAME";




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
        string colpt18 = "STATION_LINEAR";
        string colpt19 = "STATION_IFC";

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
        double length_check_tolerance = 0.5;

        System.Data.DataTable dt_display;

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



        private void transfer_errors_to_panel(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                dt_display = dt1.Copy();
                dt_display.Columns.RemoveAt(4);
                dt_display.Columns.RemoveAt(2);
                dataGridView_error_weld_map.DataSource = dt_display;
                dataGridView_error_weld_map.Columns[0].Width = 300;
                dataGridView_error_weld_map.Columns[1].Width = 75;
                dataGridView_error_weld_map.Columns[2].Width = 300;
                dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_weld_map.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_weld_map.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_error_weld_map.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_error_weld_map.EnableHeadersVisualStyles = false;
            }
        }

        private void display_errors(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                dt1.Columns.RemoveAt(6);
                dt1.Columns.RemoveAt(5);
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
                if (Wgen_main_form.dt_weld_map.Rows[i][col2] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col3] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col4] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col5] != DBNull.Value &&
                    Wgen_main_form.dt_weld_map.Rows[i][col7] != DBNull.Value &&
                    (Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToLower().Replace(" ", "") == "weld" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToLower().Replace(" ", "") == "wld" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToLower().Replace(" ", "") == "bend" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToLower().Replace(" ", "") == "elbow" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToLower().Replace(" ", "") == "bore_face" ||
                    Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col5]).ToUpper().Replace(" ", "") == "LOOSE_END") &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col2])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col3])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col4])) == true &&
                    Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col7]).Replace("+", "")) == true)
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
                            string fc1 = Convert.ToString(dt1.Rows[i][col5]).Replace(" ", "");
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

                                double x1 = Convert.ToDouble(dt1.Rows[i][col3]);
                                double y1 = Convert.ToDouble(dt1.Rows[i][col2]);
                                double z1 = Convert.ToDouble(dt1.Rows[i][col4]);

                                double x2 = 0;
                                double y2 = 0;
                                double z2 = 0;

                                string wt_coating = "wt";
                                if (dt1.Rows[i][col16] != DBNull.Value)
                                {
                                    wt_coating = Convert.ToString(dt1.Rows[i][col16]);
                                }

                                if (dt1.Rows[i][col19] != DBNull.Value)
                                {
                                    wt_coating = wt_coating + "_" + Convert.ToString(dt1.Rows[i][col19]);
                                }

                                #region OBJECT DATA
                                string MMID = "";
                                if (dt1.Rows[i][col15] != DBNull.Value) MMID = Convert.ToString(dt1.Rows[i][col15]);

                                string PIPEID = "";
                                if (dt1.Rows[i][col17] != DBNull.Value) PIPEID = Convert.ToString(dt1.Rows[i][col17]);

                                string HEAT = "";
                                if (dt1.Rows[i][col18] != DBNull.Value) HEAT = Convert.ToString(dt1.Rows[i][col18]);

                                string Xray1 = "";
                                if (dt1.Rows[i][col6] != DBNull.Value) Xray1 = Convert.ToString(dt1.Rows[i][col6]);

                                string COATING = "";
                                if (dt1.Rows[i][col19] != DBNull.Value) COATING = Convert.ToString(dt1.Rows[i][col19]);

                                string STA_START = "";
                                if (dt1.Rows[i][col7] != DBNull.Value)
                                {
                                    STA_START = Convert.ToString(dt1.Rows[i][col7]);
                                    if (Functions.IsNumeric(STA_START.Replace("+", "")) == true)
                                    {
                                        STA_START = Functions.Get_chainage_from_double(Convert.ToDouble(STA_START.Replace("+", "")), "f", 2);
                                    }
                                }

                                string WALL = "";
                                if (dt1.Rows[i][col16] != DBNull.Value) WALL = Convert.ToString(dt1.Rows[i][col16]);

                                string PNT_START = "";
                                if (dt1.Rows[i][col1] != DBNull.Value) PNT_START = Convert.ToString(dt1.Rows[i][col1]);

                                int no_bore_face = 0;
                                PolylineVertex3d[] Vertex_new_bend = new PolylineVertex3d[0];
                                if (i < dt1.Rows.Count - 1)
                                {
                                    #region j
                                    for (int j = i + 1; j < dt1.Rows.Count; ++j)
                                    {
                                        fc2 = Convert.ToString(dt1.Rows[j][col5]).Replace(" ", "");

                                        if (fc2.ToLower() == "bend" || fc2.ToLower() == "elbow")
                                        {

                                            double xb = Convert.ToDouble(dt1.Rows[j][col3]);
                                            double yb = Convert.ToDouble(dt1.Rows[j][col2]);
                                            double zb = Convert.ToDouble(dt1.Rows[j][col4]);

                                            Array.Resize(ref Vertex_new_bend, Vertex_new_bend.Length + 1);
                                            Vertex_new_bend[Vertex_new_bend.Length - 1] = new PolylineVertex3d(new Point3d(xb, yb, zb));

                                            #region Point_block

                                            System.Collections.Specialized.StringCollection Col_name1 = new System.Collections.Specialized.StringCollection();
                                            System.Collections.Specialized.StringCollection Col_value1 = new System.Collections.Specialized.StringCollection();

                                            string PNT_BEND = "";
                                            if (dt1.Rows[j][col1] != DBNull.Value) PNT_BEND = Convert.ToString(dt1.Rows[j][col1]);

                                            Col_name1.Add("PTNO");
                                            Col_value1.Add(PNT_BEND);

                                            Col_name1.Add("FEATURE_CODE");
                                            if (dt1.Rows[j][col5] != DBNull.Value) Col_value1.Add(Convert.ToString(dt1.Rows[j][col5]));

                                            string DESC1 = "";
                                            if (dt1.Rows[j][col6] != DBNull.Value) DESC1 = Convert.ToString(dt1.Rows[j][col6]);
                                            Col_name1.Add("DESCRIPTION");
                                            Col_value1.Add(DESC1);

                                            string STA_b = "";
                                            if (dt1.Rows[j][col7] != DBNull.Value)
                                            {
                                                STA_b = Convert.ToString(dt1.Rows[j][col7]);
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
                                            if (dt1.Rows[j][col1] != DBNull.Value)
                                            {
                                                if (PNT_END == "")
                                                {
                                                    PNT_END = Convert.ToString(dt1.Rows[j][col1]);
                                                }
                                                else
                                                {
                                                    PNT_END = PNT_END + ", " + Convert.ToString(dt1.Rows[j][col1]);
                                                }

                                            }

                                            if ((dt1.Rows[j][col6]) != DBNull.Value)
                                            {
                                                Xray2 = Convert.ToString(dt1.Rows[j][col6]);
                                            }
                                            #endregion

                                        }
                                        else if (fc2.ToLower() == "weld" || fc2.ToLower() == "loose_end" || fc2.ToLower() == "wld" || fc2.ToLower() == "le")
                                        {
                                            x2 = Convert.ToDouble(dt1.Rows[j][col3]);
                                            y2 = Convert.ToDouble(dt1.Rows[j][col2]);
                                            z2 = Convert.ToDouble(dt1.Rows[j][col4]);

                                            #region OBJECT DATA
                                            if (dt1.Rows[j][col1] != DBNull.Value)
                                            {
                                                if (PNT_END != "")
                                                {
                                                    PNT_END = PNT_END + ", " + Convert.ToString(dt1.Rows[j][col1]);
                                                }
                                                else
                                                {
                                                    PNT_END = Convert.ToString(dt1.Rows[j][col1]);
                                                }
                                            }

                                            if (dt1.Rows[j][col7] != DBNull.Value)
                                            {
                                                STA_END = Convert.ToString(dt1.Rows[j][col7]);
                                                if (Functions.IsNumeric(STA_END.Replace("+", "")) == true)
                                                {
                                                    STA_END = Functions.Get_chainage_from_double(Convert.ToDouble(STA_END.Replace("+", "")), "f", 2);
                                                }
                                            }

                                            if ((dt1.Rows[j][col6]) != DBNull.Value)
                                            {
                                                Xray2 = Convert.ToString(dt1.Rows[j][col6]);
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
                                    if (dt1.Rows[i][col9] != DBNull.Value)
                                    {
                                        string MM_BACK = Convert.ToString(dt1.Rows[i][col9]);
                                        Col_name.Add("MM_BACK");
                                        Col_value.Add(MM_BACK);
                                    }

                                    if (dt1.Rows[i][col15] != DBNull.Value)
                                    {
                                        string MM_AHEAD = Convert.ToString(dt1.Rows[i][col15]);
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
                                        if (dt1.Rows[k][col1] != DBNull.Value) pnt1 = Convert.ToString(dt1.Rows[k][col1]);
                                        Col_namek.Add("PTNO");
                                        Col_valuek.Add(pnt1);


                                        string fc = "";
                                        if (dt1.Rows[k][col5] != DBNull.Value) fc = Convert.ToString(dt1.Rows[k][col5]).Replace(" ", "");
                                        Col_namek.Add("FEATURE_CODE");
                                        Col_valuek.Add(fc);

                                        string sta1 = "";
                                        if (dt1.Rows[k][col7] != DBNull.Value)
                                        {
                                            sta1 = Convert.ToString(dt1.Rows[k][col7]);
                                            if (Functions.IsNumeric(sta1.Replace("+", "")) == true)
                                            {
                                                sta1 = Functions.Get_chainage_from_double(Convert.ToDouble(sta1.Replace("+", "")), "f", 2);
                                            }
                                        }
                                        Col_namek.Add("STATION");
                                        Col_valuek.Add(sta1);

                                        string descr = "";
                                        if (dt1.Rows[k][col6] != DBNull.Value) descr = Convert.ToString(dt1.Rows[k][col6]);
                                        Col_namek.Add("DESCRIPTION");
                                        Col_valuek.Add(descr);

                                        double xk = Convert.ToDouble(dt1.Rows[k][col3]);
                                        double yk = Convert.ToDouble(dt1.Rows[k][col2]);
                                        double zk = Convert.ToDouble(dt1.Rows[k][col4]);



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
        private void Add_station_to_all_points(Polyline cl_poly)
        {
            if (Wgen_main_form.dt_all_points != null || Wgen_main_form.dt_all_points.Rows.Count >= 0)
            {
                if (Wgen_main_form.dt_all_points.Columns.Contains(col7) == false)
                {
                    Wgen_main_form.dt_all_points.Columns.Add(col7, typeof(double));
                }

                if (Wgen_main_form.dt_all_points.Columns.Contains(col8) == false)
                {
                    Wgen_main_form.dt_all_points.Columns.Add(col8, typeof(double));
                }

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

                        Wgen_main_form.dt_all_points.Rows[i][col7] = sta1;
                        if (dt_st_eq != null && dt_st_eq.Rows.Count > 0)
                        {
                            double sta_eq = Station_equation_of(sta1);
                            Wgen_main_form.dt_all_points.Rows[i][col8] = sta_eq;
                        }
                    }
                }
            }
        }

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

        private void button_create_weldmap_Click(object sender, EventArgs e)
        {
            if (Wgen_main_form.dt_all_points == null || Wgen_main_form.dt_all_points.Rows.Count == 0 || Wgen_main_form.dt_ground_tally == null || Wgen_main_form.dt_ground_tally.Rows.Count == 0)
            {
                return;
            }

            if (MessageBox.Show("Select the centerline", "WGEN", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
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
                Wgen_main_form.dt_all_points = Functions.Sort_data_table(Wgen_main_form.dt_all_points, colpt18);
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

                                if ((Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "BEND") ||
                                    (Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]) == "ELBOW"))
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
                                            check15.Contains("{P}") == true
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
                                            check16.Contains("{P}") == true
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
                                            check17.Contains("{P}") == true
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
                                            check18.Contains("{P}") == true
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
                                            check19.Contains("{P}") == true
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

                DataRelation relation1 = new DataRelation("xxx", Wgen_main_form.dt_all_points.Columns[col_mm_back], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
                dataset1.Relations.Add(relation1);

                DataRelation relation2 = new DataRelation("xxx1", Wgen_main_form.dt_all_points.Columns[col_mm_ahead], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
                dataset1.Relations.Add(relation2);

                //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(Wgen_main_form.dt_all_points);

                for (i = 0; i < Wgen_main_form.dt_all_points.Rows.Count; ++i)
                {
                    if (Wgen_main_form.dt_all_points.Rows[i][colpt1] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt5] != DBNull.Value &&
                        Wgen_main_form.dt_all_points.Rows[i][colpt18] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt18]).Replace("+", "")) == true &&
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
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col5] = Feature_code.ToUpper();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col7] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col7]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col8] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col8]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col1] = pt1;
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col2] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col3] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col4] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4]);

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
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col6] = descr;
                                            j = Wgen_main_form.dt_feature_codes.Rows.Count;
                                        }
                                    }
                                    #endregion


                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt7] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col26] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt7]);
                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt6] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col27] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt6]);

                                    dt_colors.Rows.Add();

                                    int nr_match = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1).Length;
                                    if (nr_match == 1)
                                    {
                                        System.Data.DataRow row1 = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation1)[0];

                                        if (row1[colgt1] != DBNull.Value)
                                        {
                                            string mmid = Convert.ToString(row1[colgt1]);
                                            string pipeID = "";
                                            string heat = "";
                                            string wt = "";
                                            string diam = "";
                                            string grd = "";
                                            string coating = "";

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

                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col9] = mmid;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col10] = wt;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col11] = pipeID;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col12] = heat;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col13] = coating;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col14] = grd;


                                        }
                                    }

                                    if (Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation2).Length == 1)
                                    {
                                        System.Data.DataRow row1 = Wgen_main_form.dt_all_points.Rows[i].GetChildRows(relation2)[0];

                                        if (row1[colgt1] != DBNull.Value)
                                        {
                                            string mmid = Convert.ToString(row1[colgt1]);
                                            string pipeID = "";
                                            string heat = "";
                                            string wt = "";
                                            string diam = "";
                                            string grd = "";
                                            string coating = "";

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


                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col15] = mmid;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col16] = wt;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col17] = pipeID;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col18] = heat;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col19] = coating;
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col20] = grd;

                                        }
                                    }
                                }

                                else if (Wgen_main_form.lista_feature_code.Contains(Feature_code.ToUpper()) == true)
                                {
                                    Wgen_main_form.dt_weld_map.Rows.Add();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col5] = Feature_code.ToUpper();
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col7] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col7]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col8] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][col8]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col1] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt1]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col2] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt2]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col3] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt3]);
                                    Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col4] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt4]);

                                    #region FEATURE CODES MAPPING
                                    for (int j = 0; j < Wgen_main_form.dt_feature_codes.Rows.Count; ++j)
                                    {
                                        if (Wgen_main_form.client_name == Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][0]) &&
                                            Feature_code.ToUpper() == Convert.ToString(Wgen_main_form.dt_feature_codes.Rows[j][1]).ToUpper() &&
                                           (bool)Wgen_main_form.dt_feature_codes.Rows[j][2] == true)
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
                                            Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col6] = descr;
                                            j = Wgen_main_form.dt_feature_codes.Rows.Count;
                                        }
                                    }
                                    #endregion

                                    if (Feature_code.ToUpper() == "BEND" || Feature_code.ToUpper() == "ELBOW")
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
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col6] =
                                                    bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR";
                                            }
                                            else if (bend_deflection_left_right_sag_overbnd == "SAG" || bend_deflection_left_right_sag_overbnd == "OVERBEND")
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col6] =
                                                    bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + vdefl + " VER";
                                            }
                                            else
                                            {
                                                Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col6] =
                                                    bend_type_field_induction + "/" + bend_deflection_left_right_sag_overbnd + "/" + position + "/" + hdefl + " HOR" + "/" + vdefl + " VER";
                                            }
                                        }
                                    }


                                    if (Wgen_main_form.dt_all_points.Rows[i][colpt6] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[Wgen_main_form.dt_weld_map.Rows.Count - 1][col27] = Convert.ToString(Wgen_main_form.dt_all_points.Rows[i][colpt6]);

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
                                    if (Feature_code.ToUpper() == "RIVER_WEIGHT")
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


                dataset1.Relations.Remove(relation1);
                dataset1.Relations.Remove(relation2);
                dataset1.Tables.Remove(Wgen_main_form.dt_all_points);
                dataset1.Tables.Remove(Wgen_main_form.dt_ground_tally);

                if (dt_ng.Rows.Count > 0)
                {
                    //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_ng);

                    Wgen_main_form.dt_weld_map.Columns.Add("Ground Point to Weld Point Distance", typeof(double));
                    for (i = 0; i < Wgen_main_form.dt_weld_map.Rows.Count; ++i)
                    {

                        string pt_weld = Convert.ToString(Wgen_main_form.dt_weld_map.Rows[i][col1]);

                        Point3d pt1 = new Point3d(Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col3]), Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col2]), 0);

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
                        Wgen_main_form.dt_weld_map.Rows[i][col21] = ptg;
                        Wgen_main_form.dt_weld_map.Rows[i][col22] = Convert.ToString(Ymin);
                        Wgen_main_form.dt_weld_map.Rows[i][col23] = Convert.ToString(Xmin);
                        Wgen_main_form.dt_weld_map.Rows[i][col24] = Convert.ToString(Zmin);
                        if (Wgen_main_form.dt_weld_map.Rows[i][col4] != DBNull.Value) Wgen_main_form.dt_weld_map.Rows[i][col25] = Convert.ToString(Zmin - Convert.ToDouble(Wgen_main_form.dt_weld_map.Rows[i][col4]));



                    }
                }


                if (Wgen_main_form.dt_weld_map.Rows.Count > 0)
                {
                    transfer_weld_coordinates_to_pmc(Wgen_main_form.dt_weld_map);
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
            dt_welds.Columns.Add(col1, typeof(string));
            dt_welds.Columns.Add(col2, typeof(double));
            dt_welds.Columns.Add(col3, typeof(double));
            dt_welds.Columns.Add(col4, typeof(double));
            dt_welds.Columns.Add(col7, typeof(string));
            dt_welds.Columns.Add(col8, typeof(string));

            System.Data.DataTable dt_pmc = new System.Data.DataTable();
            dt_pmc = dt_welds.Clone();
            dt_pmc.Columns.Add("index", typeof(int));

            for (int i = 0; i < dtwm.Rows.Count; ++i)
            {
                string feature1 = "xx";
                if (dtwm.Rows[i][col5] != DBNull.Value)
                {
                    feature1 = Convert.ToString(dtwm.Rows[i][col5]);
                }

                if (feature1.ToUpper() == "PIPE_MATERIAL_CHANGE" || feature1.ToUpper() == "COATING_CHANGE" || feature1.ToUpper() == "PIP" || feature1.ToUpper() == "CC")
                {
                    dt_pmc.Rows.Add();
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col1] = dtwm.Rows[i][col1];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col7] = dtwm.Rows[i][col7];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1][col8] = dtwm.Rows[i][col8];
                    dt_pmc.Rows[dt_pmc.Rows.Count - 1]["index"] = i;
                }

                if (feature1.ToUpper() == "WELD" || feature1.ToUpper() == "WLD")
                {
                    dt_welds.Rows.Add();
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col1] = dtwm.Rows[i][col1];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col7] = dtwm.Rows[i][col7];
                    dt_welds.Rows[dt_welds.Rows.Count - 1][col8] = dtwm.Rows[i][col8];
                }
            }


            for (int i = 0; i < dt_pmc.Rows.Count; ++i)
            {
                double dmax = 0.5;
                if (dt_pmc.Rows[i][col2] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[i][col2])) == true &&
                    dt_pmc.Rows[i][col3] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_pmc.Rows[i][col3])) == true)
                {
                    double y1 = Convert.ToDouble(dt_pmc.Rows[i][col2]);
                    double x1 = Convert.ToDouble(dt_pmc.Rows[i][col3]);
                    for (int j = 0; j < dt_welds.Rows.Count; ++j)
                    {
                        if (dt_welds.Rows[j][col2] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds.Rows[j][col2])) == true &&
                            dt_welds.Rows[j][col3] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt_welds.Rows[j][col3])) == true)
                        {
                            double y2 = Convert.ToDouble(dt_welds.Rows[j][col2]);
                            double x2 = Convert.ToDouble(dt_welds.Rows[j][col3]);
                            double d1 = Math.Pow(Math.Pow(x1 - x2, 2) + Math.Pow(y1 - y2, 2), 0.5);
                            if (d1 < dmax)
                            {
                                dmax = d1;
                                dt_pmc.Rows[i][col2] = dt_welds.Rows[j][col2];
                                dt_pmc.Rows[i][col3] = dt_welds.Rows[j][col3];
                                dt_pmc.Rows[i][col4] = dt_welds.Rows[j][col4];
                                dt_pmc.Rows[i][col7] = dt_welds.Rows[j][col7];
                                dt_pmc.Rows[i][col8] = dt_welds.Rows[j][col8];
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < dt_pmc.Rows.Count; ++i)
            {
                int idx = Convert.ToInt32(dt_pmc.Rows[i]["index"]);
                dtwm.Rows[idx][col2] = dt_pmc.Rows[i][col2];
                dtwm.Rows[idx][col3] = dt_pmc.Rows[i][col3];
                dtwm.Rows[idx][col4] = dt_pmc.Rows[i][col4];
                dtwm.Rows[idx][col7] = dt_pmc.Rows[i][col7];
                dtwm.Rows[idx][col8] = dt_pmc.Rows[i][col8];

            }

        }

        private void wm_checks(System.Data.DataTable dtwm)
        {
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


            var duplicates_pts = dtwm.AsEnumerable().GroupBy(datarow1 => new { pnt = datarow1.Field<string>(col1) }).Where(g => g.Count() > 1).Select(g => new { g.Key.pnt }).ToList();
            var duplicates_mmid_back = dtwm.AsEnumerable().GroupBy(datarow1 => new { mmid = datarow1.Field<string>(col9) }).Where(g => g.Count() > 1).Select(g => new { g.Key.mmid }).ToList();
            var duplicates_mmid_ahead = dtwm.AsEnumerable().GroupBy(datarow1 => new { mmid = datarow1.Field<string>(col15) }).Where(g => g.Count() > 1).Select(g => new { g.Key.mmid }).ToList();

            System.Data.DataTable dt2 = new System.Data.DataTable();
            dt2.Columns.Add(col1, typeof(string));
            dt2.TableName = "dt2";

            System.Data.DataTable dt_welds = new System.Data.DataTable();
            dt_welds.Columns.Add(col1, typeof(string));
            dt_welds.Columns.Add(col2, typeof(double));
            dt_welds.Columns.Add(col3, typeof(double));
            dt_welds.Columns.Add(col4, typeof(double));
            dt_welds.Columns.Add(col7, typeof(string));

            dt_welds.Columns.Add(col10, typeof(string));
            dt_welds.Columns.Add(col16, typeof(string));
            dt_welds.Columns.Add(col13, typeof(string));
            dt_welds.Columns.Add(col19, typeof(string));
            dt_welds.Columns.Add("address", typeof(string));

            System.Data.DataTable dt_pmc = new System.Data.DataTable();
            dt_pmc = dt_welds.Clone();
            dt_pmc.Columns.Add(col6, typeof(string));

            System.Data.DataTable dt_cc = new System.Data.DataTable();
            dt_cc = dt_welds.Clone();

            System.Data.DataTable dt_bend_welds = new System.Data.DataTable();
            dt_bend_welds.Columns.Add(col1, typeof(string));
            dt_bend_welds.Columns.Add(col2, typeof(double));
            dt_bend_welds.Columns.Add(col3, typeof(double));
            dt_bend_welds.Columns.Add(col4, typeof(double));
            dt_bend_welds.Columns.Add(col15, typeof(string));
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
                DataRelation relation1 = new DataRelation("xxx1", dtwm.Columns[col1], dt2.Columns[col1], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate points
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col5] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col5]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string Point1 = dtwm.Rows[i].GetChildRows(relation1)[j][col1].ToString();

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
            dt2.Columns.Add(col9, typeof(string));
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
                DataRelation relation1 = new DataRelation("xxx9", dtwm.Columns[col9], dt2.Columns[col9], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = nr_duplicates + dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate mmid
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string pt1 = "xx";
                        if (dtwm.Rows[i][col1] != DBNull.Value)
                        {
                            pt1 = Convert.ToString(dtwm.Rows[i][col1]);
                        }
                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col5] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col5]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string mmid1 = dtwm.Rows[i].GetChildRows(relation1)[j][col9].ToString();

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
            dt2.Columns.Add(col15, typeof(string));
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
                DataRelation relation1 = new DataRelation("xxx9", dtwm.Columns[col15], dt2.Columns[col15], false);

                dataset1.Relations.Add(relation1);

                nr_duplicates = nr_duplicates + dt2.Rows.Count;

                for (int i = 0; i < dtwm.Rows.Count; ++i)
                {
                    #region duplicate mmid
                    if (dtwm.Rows[i].GetChildRows(relation1).Length > 0)
                    {
                        string pt1 = "xx";
                        if (dtwm.Rows[i][col1] != DBNull.Value)
                        {
                            pt1 = Convert.ToString(dtwm.Rows[i][col1]);
                        }

                        string Feature1 = "xx";
                        if (dtwm.Rows[i][col5] != DBNull.Value)
                        {
                            Feature1 = Convert.ToString(dtwm.Rows[i][col5]);
                        }

                        string x = "";
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        for (int j = 0; j < dtwm.Rows[i].GetChildRows(relation1).Length; ++j)
                        {
                            string mmid1 = dtwm.Rows[i].GetChildRows(relation1)[j][col15].ToString();

                            dt_errors.Rows.Add();
                            dt_errors.Rows[dt_errors.Rows.Count - 1][0] = pt1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][1] = Feature1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][2] = mmid1;
                            dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_15.Text + Convert.ToString(i + start_row);
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



            string status_rock_shield = "END";
            for (int i = 0; i < dtwm.Rows.Count; ++i)
            {
                string feature_rs1 = "xx";
                if (dtwm.Rows[i][col5] != DBNull.Value)
                {
                    feature_rs1 = Convert.ToString(dtwm.Rows[i][col5]);
                }

                #region null values
                if (dtwm.Rows[i][col1] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col1]) == "")
                {
                    dt_errors.Rows.Add();
                    dt_errors.Rows[dt_errors.Rows.Count - 1][1] = feature_rs1;
                    dt_errors.Rows[dt_errors.Rows.Count - 1][3] = textBox_1.Text + Convert.ToString(i + start_row);
                    dt_errors.Rows[dt_errors.Rows.Count - 1][4] = "No Point ID Specified";
                    ++nr_null_values;
                    lista_puncte.Add("null");
                }

                if (dtwm.Rows[i][col1] != DBNull.Value)
                {
                    string pt1 = Convert.ToString(dtwm.Rows[i][col1]).ToUpper();
                    int index1 = -1;



                    if (dtwm.Rows[i][col2] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col2]) == "")
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
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col2])) == false)
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
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col2]);
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
                    if (dtwm.Rows[i][col3] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col3]) == "")
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
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col3])) == false)
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
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col3]);
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

                    if (dtwm.Rows[i][col4] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col4]) == "")
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
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;
                        ++nr_null_values;
                    }
                    else
                    {
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col4])) == false)
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
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col4]);
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
                            if (dtwm.Rows[i][col3] != DBNull.Value)
                            {
                                x = Convert.ToString(dtwm.Rows[i][col3]);
                            }
                            string y = "";
                            if (dtwm.Rows[i][col2] != DBNull.Value)
                            {
                                y = Convert.ToString(dtwm.Rows[i][col2]);
                            }

                            dt_errors.Rows[index1][5] = x;
                            dt_errors.Rows[index1][6] = y;
                            ++nr_null_values;
                        }
                    }
                    if (dtwm.Rows[i][col5] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col5]) == "")
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
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }

                    if (dtwm.Rows[i][col6] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col6]) == "")
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
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }

                    if (dtwm.Rows[i][col7] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col7]) == "")
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
                        if (dtwm.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dtwm.Rows[i][col3]);
                        }
                        string y = "";
                        if (dtwm.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dtwm.Rows[i][col2]);
                        }

                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;

                        ++nr_null_values;
                    }
                    else
                    {
                        bool adauga = false;
                        if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col7])) == false)
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
                            dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col7]);
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
                            if (dtwm.Rows[i][col3] != DBNull.Value)
                            {
                                x = Convert.ToString(dtwm.Rows[i][col3]);
                            }
                            string y = "";
                            if (dtwm.Rows[i][col2] != DBNull.Value)
                            {
                                y = Convert.ToString(dtwm.Rows[i][col2]);
                            }

                            dt_errors.Rows[index1][5] = x;
                            dt_errors.Rows[index1][6] = y;

                            ++nr_null_values;
                        }
                    }

                    if (dtwm.Rows[i][col5] != DBNull.Value)
                    {
                        if (Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "WLD")
                        {
                            bool is_fab1 = false;
                            if (dtwm.Rows[i][col9] != DBNull.Value)
                            {
                                if (Convert.ToString(dtwm.Rows[i][col9]).ToUpper() == "FAB")
                                {
                                    is_fab1 = true;
                                }
                            }

                            bool is_fab2 = false;
                            if (dtwm.Rows[i][col15] != DBNull.Value)
                            {
                                if (Convert.ToString(dtwm.Rows[i][col15]).ToUpper() == "FAB")
                                {
                                    is_fab2 = true;
                                }
                            }

                            if ((dtwm.Rows[i][col10] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col10]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Wall Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Wall Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col10])) == false && is_fab1 == false)
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
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col10]);
                                    dt_errors.Rows[index1][3] = textBox_10.Text + Convert.ToString(i + start_row);
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
                                    if (dtwm.Rows[i][col3] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col3]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col2] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col2]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }
                            if ((dtwm.Rows[i][col9] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col9]) == "") && is_fab1 == false)
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
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col11] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col11]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Pipe ID Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe ID Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col12] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col12]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Heat Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Heat Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;

                            }
                            if ((dtwm.Rows[i][col13] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col13]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Coating Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Coating Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col14] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col14]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Pipe Grade Back Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe Grade Back Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if (dtwm.Rows[i][col15] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col15]) == "")
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
                                    dt_errors.Rows[index1][4] = "No MMID Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No MMID Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col16] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col16]) == "") && is_fab1 == false)
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
                                    dt_errors.Rows[index1][4] = "No Wall Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Wall Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            else
                            {
                                if (Functions.IsNumeric(Convert.ToString(dtwm.Rows[i][col16])) == false && is_fab1 == false)
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
                                    dt_errors.Rows[index1][2] = Convert.ToString(dtwm.Rows[i][col16]);
                                    dt_errors.Rows[index1][3] = textBox_16.Text + Convert.ToString(i + start_row);

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
                                    if (dtwm.Rows[i][col3] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col3]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col2] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col2]);
                                    }

                                    dt_errors.Rows[index1][5] = x;
                                    dt_errors.Rows[index1][6] = y;

                                    ++nr_null_values;
                                }
                            }
                            if ((dtwm.Rows[i][col17] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col17]) == "") && is_fab2 == false)
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
                                    dt_errors.Rows[index1][4] = "No Pipe ID Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Pipe ID Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col18] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col18]) == "") && is_fab2 == false)
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
                                    dt_errors.Rows[index1][4] = "No Heat Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Heat Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col19] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col19]) == "") && is_fab2 == false)
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
                                    dt_errors.Rows[index1][4] = "No Coating Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Coating Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;


                                ++nr_null_values;
                            }
                            if ((dtwm.Rows[i][col20] == DBNull.Value || Convert.ToString(dtwm.Rows[i][col20]) == "") && is_fab2 == false)
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
                                    dt_errors.Rows[index1][4] = "No Grade Pipe Ahead Specified";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "No Grade Pipe Ahead Specified";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }

                                string x = "";
                                if (dtwm.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dtwm.Rows[i][col3]);
                                }
                                string y = "";
                                if (dtwm.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dtwm.Rows[i][col2]);
                                }

                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                                ++nr_null_values;
                            }
                        }
                    }
                }
                #endregion


                if (feature_rs1 != "xx")
                {
                    if (dtwm.Rows[i][col1] != DBNull.Value)
                    {
                        string pt1 = Convert.ToString(dtwm.Rows[i][col1]).ToUpper();
                        #region BEND
                        if (Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "BEND" || Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "ELBOW")
                        {
                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col1] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col15] = dtwm.Rows[i][col15];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "B";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion

                        #region loose end
                        if (Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "LOOSE_END" || Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "LE")
                        {
                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col1] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col15] = dtwm.Rows[i][col15];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "X";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion

                        #region WELD
                        if (Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[i][col5]).ToUpper() == "WLD")
                        {

                            dt_welds.Rows.Add();
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col1] = pt1;
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col7] = dtwm.Rows[i][col7];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col10] = dtwm.Rows[i][col10];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col16] = dtwm.Rows[i][col16];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col13] = dtwm.Rows[i][col13];
                            dt_welds.Rows[dt_welds.Rows.Count - 1][col19] = dtwm.Rows[i][col19];
                            dt_welds.Rows[dt_welds.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);


                            dt_bend_welds.Rows.Add();
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col1] = pt1;
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1][col15] = dtwm.Rows[i][col15];
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["feature_code"] = "W";
                            dt_bend_welds.Rows[dt_bend_welds.Rows.Count - 1]["address"] = textBox_15.Text + Convert.ToString(i + start_row);


                            bool is_loose_end = false;
                            for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                            {
                                if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "LOOSE_END" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "LE")
                                {
                                    is_loose_end = true;
                                    j = dtwm.Rows.Count;
                                }
                                if (is_loose_end == false && (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD"))
                                {
                                    j = dtwm.Rows.Count;
                                }
                            }

                            if (is_loose_end == false)
                            {

                                int index1 = -1;

                                #region mmid back-ahead
                                if (dtwm.Rows[i][col15] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col15]) != "")
                                {
                                    string MM1 = Convert.ToString(dtwm.Rows[i][col15]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col9] != DBNull.Value)
                                                {
                                                    string MM2 = Convert.ToString(dtwm.Rows[j][col9]);
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
                                                        dt_errors.Rows[index1][3] = textBox_15.Text + Convert.ToString(i + start_row);

                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "MM id Ahead vs Back Missmatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "MM id Ahead vs Back Missmatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }

                                                        string x = "";
                                                        if (dtwm.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col2]);
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
                                if (dtwm.Rows[i][col16] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col16]) != "")
                                {
                                    string wall1 = Convert.ToString(dtwm.Rows[i][col16]);

                                    if (i < dtwm.Rows.Count - 1 && Functions.IsNumeric(wall1) == true)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            string wt_change_descr = "";
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "COATING_CHANGE")
                                            {
                                                if (dtwm.Rows[j][col6] != DBNull.Value)
                                                {
                                                    wt_change_descr = Convert.ToString(dtwm.Rows[j][col6]);
                                                }
                                            }
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "PIPE_MATERIAL_CHANGE")
                                            {
                                                if (dtwm.Rows[j][col6] != DBNull.Value)
                                                {
                                                    wt_change_descr = Convert.ToString(dtwm.Rows[j][col6]);
                                                }
                                            }
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col10] != DBNull.Value)
                                                {
                                                    string wall2 = Convert.ToString(dtwm.Rows[j][col10]);
                                                    if (wt_change_descr != "")
                                                    {
                                                        if ((wall1 + "TO" + wall2).ToUpper().Replace(" ", "").ToUpper() != wt_change_descr)
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
                                                            dt_errors.Rows[index1][2] = "Wall: " + wall1 + " vs. " + wall2;
                                                            dt_errors.Rows[index1][3] = textBox_16.Text + Convert.ToString(i + start_row);

                                                            if (adauga == true)
                                                            {
                                                                dt_errors.Rows[index1][4] = "Wall Ahead vs Back Missmatch";
                                                            }
                                                            else
                                                            {
                                                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Wall Ahead vs Back Missmatch";
                                                                dt_errors.Rows[index1][4] = Existing_error;
                                                            }

                                                            string x = "";
                                                            if (dtwm.Rows[i][col3] != DBNull.Value)
                                                            {
                                                                x = Convert.ToString(dtwm.Rows[i][col3]);
                                                            }
                                                            string y = "";
                                                            if (dtwm.Rows[i][col2] != DBNull.Value)
                                                            {
                                                                y = Convert.ToString(dtwm.Rows[i][col2]);
                                                            }

                                                            dt_errors.Rows[index1][5] = x;
                                                            dt_errors.Rows[index1][6] = y;

                                                        }
                                                    }
                                                    else if (Functions.IsNumeric(wall2) == true)
                                                    {
                                                        if (wall1.ToUpper() != wall2.ToUpper())
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
                                                            dt_errors.Rows[index1][2] = "Wall: " + wall1 + " vs. " + wall2;
                                                            dt_errors.Rows[index1][3] = textBox_16.Text + Convert.ToString(i + start_row);
                                                            if (adauga == true)
                                                            {
                                                                dt_errors.Rows[index1][4] = "Wall Ahead vs Back Missmatch";
                                                            }
                                                            else
                                                            {
                                                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Wall Ahead vs Back Missmatch";
                                                                dt_errors.Rows[index1][4] = Existing_error;
                                                            }
                                                            string x = "";
                                                            if (dtwm.Rows[i][col3] != DBNull.Value)
                                                            {
                                                                x = Convert.ToString(dtwm.Rows[i][col3]);
                                                            }
                                                            string y = "";
                                                            if (dtwm.Rows[i][col2] != DBNull.Value)
                                                            {
                                                                y = Convert.ToString(dtwm.Rows[i][col2]);
                                                            }

                                                            dt_errors.Rows[index1][5] = x;
                                                            dt_errors.Rows[index1][6] = y;
                                                        }
                                                    }
                                                }

                                                j = dtwm.Rows.Count;
                                            }
                                        }
                                    }
                                }
                                #endregion

                                #region pipe back-ahead
                                if (dtwm.Rows[i][col17] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col17]) != "")
                                {
                                    string pipeid1 = Convert.ToString(dtwm.Rows[i][col17]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col11] != DBNull.Value)
                                                {
                                                    string pipeid2 = Convert.ToString(dtwm.Rows[j][col11]);

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
                                                        dt_errors.Rows[index1][3] = textBox_17.Text + Convert.ToString(i + start_row);

                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Pipe id Ahead vs Back Missmatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Pipe id Ahead vs Back Missmatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }

                                                        string x = "";
                                                        if (dtwm.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col2]);
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
                                if (dtwm.Rows[i][col18] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col18]) != "")
                                {
                                    string heat1 = Convert.ToString(dtwm.Rows[i][col18]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col12] != DBNull.Value)
                                                {
                                                    string heat2 = Convert.ToString(dtwm.Rows[j][col12]);
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
                                                        dt_errors.Rows[index1][3] = textBox_18.Text + Convert.ToString(i + start_row);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Heat# Ahead vs Back Missmatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Heat# Ahead vs Back Missmatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }
                                                        string x = "";
                                                        if (dtwm.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col2]);
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

                                #region coating back-ahead
                                if (dtwm.Rows[i][col19] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col19]) != "")
                                {
                                    string coat1 = Convert.ToString(dtwm.Rows[i][col19]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            string coating_change_descr = "";
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "COATING_CHANGE")
                                            {
                                                if (dtwm.Rows[j][col6] != DBNull.Value)
                                                {
                                                    coating_change_descr = Convert.ToString(dtwm.Rows[j][col6]);
                                                }
                                            }

                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "PIPE_MATERIAL_CHANGE")
                                            {
                                                if (dtwm.Rows[j][col6] != DBNull.Value)
                                                {
                                                    coating_change_descr = Convert.ToString(dtwm.Rows[j][col6]);
                                                }
                                            }

                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col13] != DBNull.Value)
                                                {
                                                    string coat2 = Convert.ToString(dtwm.Rows[j][col13]);
                                                    if (coating_change_descr != "")
                                                    {
                                                        if ((coat1 + "TO" + coat2).ToUpper().Replace(" ", "").ToUpper() != coating_change_descr)
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
                                                            dt_errors.Rows[index1][2] = "Coating: " + coat1 + " vs. " + coat2;
                                                            dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);
                                                            if (adauga == true)
                                                            {
                                                                dt_errors.Rows[index1][4] = "Coating Ahead vs Back Missmatch";
                                                            }
                                                            else
                                                            {

                                                                string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating Ahead vs Back Missmatch";
                                                                dt_errors.Rows[index1][4] = Existing_error;
                                                            }

                                                            string x = "";
                                                            if (dtwm.Rows[i][col3] != DBNull.Value)
                                                            {
                                                                x = Convert.ToString(dtwm.Rows[i][col3]);
                                                            }
                                                            string y = "";
                                                            if (dtwm.Rows[i][col2] != DBNull.Value)
                                                            {
                                                                y = Convert.ToString(dtwm.Rows[i][col2]);
                                                            }

                                                            dt_errors.Rows[index1][5] = x;
                                                            dt_errors.Rows[index1][6] = y;

                                                        }
                                                    }
                                                    else if (coat1.ToUpper() != coat2.ToUpper())
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
                                                        dt_errors.Rows[index1][2] = "Coating: " + coat1 + " vs. " + coat2;
                                                        dt_errors.Rows[index1][3] = textBox_19.Text + Convert.ToString(i + start_row);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Coating Ahead vs Back Missmatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating Ahead vs Back Missmatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }
                                                        string x = "";
                                                        if (dtwm.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col2]);
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

                                #region grade back-ahead
                                if (dtwm.Rows[i][col20] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col20]) != "")
                                {
                                    string grade1 = Convert.ToString(dtwm.Rows[i][col20]);
                                    if (i < dtwm.Rows.Count - 1)
                                    {
                                        for (int j = i + 1; j < dtwm.Rows.Count; ++j)
                                        {
                                            if (Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WELD" || Convert.ToString(dtwm.Rows[j][col5]).ToUpper() == "WLD")
                                            {
                                                if (dtwm.Rows[j][col14] != DBNull.Value)
                                                {
                                                    string grade2 = Convert.ToString(dtwm.Rows[j][col14]);
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
                                                        dt_errors.Rows[index1][3] = textBox_20.Text + Convert.ToString(i + start_row);
                                                        if (adauga == true)
                                                        {
                                                            dt_errors.Rows[index1][4] = "Pipe Grade Ahead vs Back Missmatch";
                                                        }
                                                        else
                                                        {
                                                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Pipe Grade Ahead vs Back Missmatch";
                                                            dt_errors.Rows[index1][4] = Existing_error;
                                                        }
                                                        string x = "";
                                                        if (dtwm.Rows[i][col3] != DBNull.Value)
                                                        {
                                                            x = Convert.ToString(dtwm.Rows[i][col3]);
                                                        }
                                                        string y = "";
                                                        if (dtwm.Rows[i][col2] != DBNull.Value)
                                                        {
                                                            y = Convert.ToString(dtwm.Rows[i][col2]);
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
                            if (dtwm.Rows[i][col6] != DBNull.Value && Convert.ToString(dtwm.Rows[i][col6]) != "" && dtwm.Rows[i][col2] != DBNull.Value && dtwm.Rows[i][col3] != DBNull.Value)
                            {
                                string descr = Convert.ToString(dtwm.Rows[i][col6]);
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
                                        dt_errors.Rows[index1][4] = "Rock Shield Start/End Missmatch";
                                    }
                                    else
                                    {
                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Rock Shield Start/End Missmatch";
                                        dt_errors.Rows[index1][4] = Existing_error;
                                    }
                                    string x = "";
                                    if (dtwm.Rows[i][col3] != DBNull.Value)
                                    {
                                        x = Convert.ToString(dtwm.Rows[i][col3]);
                                    }
                                    string y = "";
                                    if (dtwm.Rows[i][col2] != DBNull.Value)
                                    {
                                        y = Convert.ToString(dtwm.Rows[i][col2]);
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

                        #region PIPE MATERIAL CHANGE
                        if (feature_rs1.ToUpper() == "PIPE_MATERIAL_CHANGE" || feature_rs1.ToUpper() == "PIP")
                        {
                            dt_pmc.Rows.Add();
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col1] = pt1;
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col7] = dtwm.Rows[i][col7];
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1][col6] = dtwm.Rows[i][col6];
                            dt_pmc.Rows[dt_pmc.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion

                        #region COATING CHANGE
                        if (feature_rs1.ToUpper() == "COATING_CHANGE" || feature_rs1.ToUpper() == "CC")
                        {
                            dt_cc.Rows.Add();
                            dt_cc.Rows[dt_cc.Rows.Count - 1][col1] = pt1;
                            dt_cc.Rows[dt_cc.Rows.Count - 1][col2] = dtwm.Rows[i][col2];
                            dt_cc.Rows[dt_cc.Rows.Count - 1][col3] = dtwm.Rows[i][col3];
                            dt_cc.Rows[dt_cc.Rows.Count - 1][col4] = dtwm.Rows[i][col4];
                            dt_cc.Rows[dt_cc.Rows.Count - 1][col7] = dtwm.Rows[i][col7];
                            dt_cc.Rows[dt_cc.Rows.Count - 1]["address"] = textBox_1.Text + Convert.ToString(i + start_row);
                        }
                        #endregion
                    }
                }
            }

            //Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_errors);

            List<string> list_of_pmc = new List<string>();
            List<string> list_of_errors = new List<string>();
            List<string> list_of_pmc_descriptions = new List<string>();


            #region PIPE MATERIAL CHANGE
            string feature1 = "PIPE_MATERIAL_CHANGE";

            if (dt_welds.Rows.Count > 0 && dt_pmc.Rows.Count > 0)
            {
                dataset1.Tables.Add(dt_pmc);
                dataset1.Tables.Add(dt_welds);
                DataRelation relation2y = new DataRelation("xxx2", dt_pmc.Columns[col2], dt_welds.Columns[col2], false);
                DataRelation relation2x = new DataRelation("xxx3", dt_pmc.Columns[col3], dt_welds.Columns[col3], false);
                DataRelation relation2z = new DataRelation("xxx4", dt_pmc.Columns[col4], dt_welds.Columns[col4], false);
                DataRelation relation_sta = new DataRelation("xxx7", dt_welds.Columns[col7], dt_pmc.Columns[col7], false);
                dataset1.Relations.Add(relation2y);
                dataset1.Relations.Add(relation2x);
                dataset1.Relations.Add(relation2z);
                dataset1.Relations.Add(relation_sta);



                for (int i = 0; i < dt_pmc.Rows.Count; ++i)
                {
                    int nr_y = dt_pmc.Rows[i].GetChildRows(relation2y).Length;
                    int nr_x = dt_pmc.Rows[i].GetChildRows(relation2x).Length;
                    int nr_z = dt_pmc.Rows[i].GetChildRows(relation2z).Length;
                    string pt1 = Convert.ToString(dt_pmc.Rows[i][col1]);


                    if (nr_y > 0 && nr_x > 0 && nr_z > 0)
                    {

                    }
                    else
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
                        dt_errors.Rows[index1][0] = pt1;


                        dt_errors.Rows[index1][1] = feature1;
                        dt_errors.Rows[index1][3] = Convert.ToString(dt_pmc.Rows[i]["address"]);

                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "X, Y, Z coordinates don't match an existing weld";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "X, Y, Z coordinates don't match an existing weld";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }
                        string x = "";
                        if (dt_pmc.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dt_pmc.Rows[i][col3]);
                        }
                        string y = "";
                        if (dt_pmc.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dt_pmc.Rows[i][col2]);
                        }
                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;
                    }
                }

                for (int i = 0; i < dt_welds.Rows.Count; ++i)
                {
                    if (dt_welds.Rows[i][col10] != DBNull.Value && dt_welds.Rows[i][col16] != DBNull.Value)
                    {
                        string wall_back = Convert.ToString(dt_welds.Rows[i][col10]);
                        string wall_ahead = Convert.ToString(dt_welds.Rows[i][col16]);
                        string pt1 = Convert.ToString(dt_welds.Rows[i][col1]);


                        if (wall_back.Replace(" ", "") != wall_ahead.Replace(" ", ""))
                        {
                            int nr_kids = dt_welds.Rows[i].GetChildRows(relation_sta).Length;
                            if (nr_kids == 0)
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

                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature1;
                                dt_errors.Rows[index1][3] = Convert.ToString(dt_welds.Rows[i]["address"]);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "PMC feature code associated with the WELD not found";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "PMC feature code associated with the WELD not found";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }
                                string x = "";
                                if (dt_welds.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dt_welds.Rows[i][col3]);
                                }
                                string y = "";
                                if (dt_welds.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dt_welds.Rows[i][col2]);
                                }
                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;
                            }
                            else
                            {
                                string descr1 = Convert.ToString(dt_welds.Rows[i].GetChildRows(relation_sta)[0][col6]);
                                list_of_pmc.Add(pt1);
                                list_of_pmc_descriptions.Add(descr1);
                            }
                        }

                    }
                }

                dataset1.Relations.Remove(relation2x);
                dataset1.Relations.Remove(relation2y);
                dataset1.Relations.Remove(relation2z);
                dataset1.Relations.Remove(relation_sta);
                dataset1.Tables.Remove(dt_pmc);
                dataset1.Tables.Remove(dt_welds);

            }


            if (dt_welds.Rows.Count > 0)
            {
                for (int i = 0; i < dt_welds.Rows.Count; ++i)
                {
                    if (dt_welds.Rows[i][col10] != DBNull.Value && dt_welds.Rows[i][col16] != DBNull.Value)
                    {
                        string wall_back = Convert.ToString(dt_welds.Rows[i][col10]);
                        string wall_ahead = Convert.ToString(dt_welds.Rows[i][col16]);
                        string pt1 = Convert.ToString(dt_welds.Rows[i][col1]);
                        if (list_of_pmc.Contains(pt1) == false)
                        {
                            if (wall_back.Replace(" ", "") != wall_ahead.Replace(" ", ""))
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
                                dt_errors.Rows[index1][0] = pt1;

                                dt_errors.Rows[index1][1] = feature1;
                                dt_errors.Rows[index1][3] = Convert.ToString(dt_welds.Rows[i]["address"]);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "PMC missmatch between back and ahead";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "PMC missmatch between back and ahead";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }
                                string x = "";
                                if (dt_welds.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dt_welds.Rows[i][col3]);
                                }
                                string y = "";
                                if (dt_welds.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dt_welds.Rows[i][col2]);
                                }
                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;
                            }
                        }
                    }
                }
            }
            #endregion

            #region COATING CHANGE

            feature1 = "COATING_CHANGE";

            if (dt_welds.Rows.Count > 0 && dt_cc.Rows.Count > 0)
            {
                dataset1.Tables.Add(dt_cc);
                dataset1.Tables.Add(dt_welds);

                DataRelation relation2y = new DataRelation("xxx2", dt_cc.Columns[col2], dt_welds.Columns[col2], false);
                DataRelation relation2x = new DataRelation("xxx3", dt_cc.Columns[col3], dt_welds.Columns[col3], false);
                DataRelation relation2z = new DataRelation("xxx4", dt_cc.Columns[col4], dt_welds.Columns[col4], false);
                DataRelation relation_sta = new DataRelation("xxx7", dt_welds.Columns[col7], dt_cc.Columns[col7], false);

                dataset1.Relations.Add(relation2y);
                dataset1.Relations.Add(relation2x);
                dataset1.Relations.Add(relation2z);
                dataset1.Relations.Add(relation_sta);



                for (int i = 0; i < dt_cc.Rows.Count; ++i)
                {
                    int nr_y = dt_cc.Rows[i].GetChildRows(relation2y).Length;
                    int nr_x = dt_cc.Rows[i].GetChildRows(relation2x).Length;
                    int nr_z = dt_cc.Rows[i].GetChildRows(relation2z).Length;
                    string pt1 = Convert.ToString(dt_cc.Rows[i][col1]);

                    if (nr_y > 0 && nr_x > 0 && nr_z > 0)
                    {

                    }
                    else
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
                        dt_errors.Rows[index1][0] = pt1;

                        dt_errors.Rows[index1][1] = feature1;
                        dt_errors.Rows[index1][3] = Convert.ToString(dt_cc.Rows[i]["address"]);
                        if (adauga == true)
                        {
                            dt_errors.Rows[index1][4] = "X, Y, Z coordinates don't match an existing weld";
                        }
                        else
                        {
                            string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "X, Y, Z coordinates don't match an existing weld";
                            dt_errors.Rows[index1][4] = Existing_error;
                        }
                        string x = "";
                        if (dt_cc.Rows[i][col3] != DBNull.Value)
                        {
                            x = Convert.ToString(dt_cc.Rows[i][col3]);
                        }
                        string y = "";
                        if (dt_cc.Rows[i][col2] != DBNull.Value)
                        {
                            y = Convert.ToString(dt_cc.Rows[i][col2]);
                        }
                        dt_errors.Rows[index1][5] = x;
                        dt_errors.Rows[index1][6] = y;
                    }



                }
                for (int i = 0; i < dt_welds.Rows.Count; ++i)
                {
                    if (dt_welds.Rows[i][col13] != DBNull.Value && dt_welds.Rows[i][col19] != DBNull.Value)
                    {
                        string coat_back = Convert.ToString(dt_welds.Rows[i][col13]);
                        string coat_ahead = Convert.ToString(dt_welds.Rows[i][col19]);
                        string pt1 = Convert.ToString(dt_welds.Rows[i][col1]);


                        bool is_listed_as_pmc = false;

                        if (list_of_pmc.Contains(pt1) == true)
                        {
                            string descr1 = list_of_pmc[list_of_pmc_descriptions.IndexOf(pt1)].ToLower();
                            if (descr1.Contains(coat_back.ToLower()) == true && descr1.Contains(coat_ahead.ToLower()) == true)
                            {
                                is_listed_as_pmc = true;
                            }
                        }

                        if (coat_back.Replace(" ", "") != coat_ahead.Replace(" ", "") && is_listed_as_pmc == false)
                        {
                            int nr_sta = dt_welds.Rows[i].GetChildRows(relation_sta).Length;
                            if (nr_sta == 0)
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
                                dt_errors.Rows[index1][0] = pt1;
                                list_of_errors.Add(pt1);
                                dt_errors.Rows[index1][1] = feature1;
                                dt_errors.Rows[index1][3] = Convert.ToString(dt_welds.Rows[i]["address"]);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "CC feature code associated with the WELD not found";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "CC feature code associated with the WELD not found";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }
                                string x = "";
                                if (dt_welds.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dt_welds.Rows[i][col3]);
                                }
                                string y = "";
                                if (dt_welds.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dt_welds.Rows[i][col2]);
                                }
                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;
                            }
                            else
                            {

                            }
                        }



                    }
                }


                dataset1.Relations.Remove(relation2x);
                dataset1.Relations.Remove(relation2y);
                dataset1.Relations.Remove(relation2z);
                dataset1.Relations.Remove(relation_sta);
                dataset1.Tables.Remove(dt_cc);
                dataset1.Tables.Remove(dt_welds);
            }

            if (dt_welds.Rows.Count > 0)
            {
                for (int i = 0; i < dt_welds.Rows.Count; ++i)
                {
                    if (dt_welds.Rows[i][col13] != DBNull.Value && dt_welds.Rows[i][col19] != DBNull.Value)
                    {
                        string coat_back = Convert.ToString(dt_welds.Rows[i][col13]);
                        string coat_ahead = Convert.ToString(dt_welds.Rows[i][col19]);
                        string pt1 = Convert.ToString(dt_welds.Rows[i][col1]);




                        if (list_of_errors.Contains(pt1) == false)
                        {
                            bool is_listed_as_pmc = false;

                            if (list_of_pmc.Contains(pt1) == true)
                            {
                                string descr1 = list_of_pmc_descriptions[list_of_pmc.IndexOf(pt1)].ToLower();
                                if (descr1.Contains(coat_back.ToLower()) == true && descr1.Contains(coat_ahead.ToLower()) == true)
                                {
                                    is_listed_as_pmc = true;
                                }
                            }


                            if (coat_back.Replace(" ", "") != coat_ahead.Replace(" ", "") && is_listed_as_pmc == false)
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
                                dt_errors.Rows[index1][0] = pt1;
                                dt_errors.Rows[index1][1] = feature1;
                                dt_errors.Rows[index1][3] = Convert.ToString(dt_welds.Rows[i]["address"]);

                                if (adauga == true)
                                {
                                    dt_errors.Rows[index1][4] = "Coating missmatch between back and ahead";
                                }
                                else
                                {
                                    string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Coating missmatch between back and ahead";
                                    dt_errors.Rows[index1][4] = Existing_error;
                                }
                                string x = "";
                                if (dt_welds.Rows[i][col3] != DBNull.Value)
                                {
                                    x = Convert.ToString(dt_welds.Rows[i][col3]);
                                }
                                string y = "";
                                if (dt_welds.Rows[i][col2] != DBNull.Value)
                                {
                                    y = Convert.ToString(dt_welds.Rows[i][col2]);
                                }
                                dt_errors.Rows[index1][5] = x;
                                dt_errors.Rows[index1][6] = y;

                            }
                        }

                    }
                }
            }


            #endregion

            if (dt_bend_welds.Rows.Count > 0)
            {
                #region lengths checks
                if (Wgen_main_form.dt_ground_tally != null && Wgen_main_form.dt_ground_tally.Rows.Count > 0)
                {
                    dataset1.Tables.Add(dt_bend_welds);
                    dataset1.Tables.Add(Wgen_main_form.dt_ground_tally);

                    DataRelation relation_pt1 = new DataRelation("xxx8", dt_bend_welds.Columns[col15], Wgen_main_form.dt_ground_tally.Columns[colgt1], false);
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
                        if (dt_bend_welds.Rows[i][col1] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col2] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col3] != DBNull.Value &&
                            dt_bend_welds.Rows[i][col4] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col1] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col2] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col3] != DBNull.Value &&
                            dt_bend_welds.Rows[prev_i][col4] != DBNull.Value)
                        {
                            double x1 = Convert.ToDouble(dt_bend_welds.Rows[i][col3]);
                            double y1 = Convert.ToDouble(dt_bend_welds.Rows[i][col2]);
                            double z1 = Convert.ToDouble(dt_bend_welds.Rows[i][col4]);
                            string pt1 = Convert.ToString(dt_bend_welds.Rows[i][col1]);
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

                                                    string pt2 = Convert.ToString(dt_bend_welds.Rows[prev_i][col1]);
                                                    string MM_ahead = "line 4794 - wm_checks";
                                                    if (dt_bend_welds.Rows[prev_i][col15] != DBNull.Value)
                                                    {
                                                        MM_ahead = Convert.ToString(dt_bend_welds.Rows[prev_i][col15]);
                                                    }


                                                    dt_errors.Rows[index1][0] = pt2 + "-" + pt1;
                                                    dt_errors.Rows[index1][1] = "WELD TO WELD";
                                                    dt_errors.Rows[index1][2] = "MM: " + MM_ahead + " - Ground Tally: " + Convert.ToString(Math.Round(d2, 2)) + " vs. calc: " + Convert.ToString(Math.Round(d1, 2));
                                                    dt_errors.Rows[index1][3] = Convert.ToString(dt_bend_welds.Rows[prev_i]["address"]);


                                                    if (adauga == true)
                                                    {
                                                        dt_errors.Rows[index1][4] = "Length missmatch with ground tally";
                                                    }
                                                    else
                                                    {
                                                        string Existing_error = Convert.ToString(dt_errors.Rows[index1][4]) + ", " + "Length missmatch with ground tally";
                                                        dt_errors.Rows[index1][4] = Existing_error;
                                                    }

                                                    string x = "";
                                                    if (dt_bend_welds.Rows[prev_i][col3] != DBNull.Value)
                                                    {
                                                        x = Convert.ToString(dt_bend_welds.Rows[prev_i][col3]);
                                                    }
                                                    string y = "";
                                                    if (dt_bend_welds.Rows[prev_i][col2] != DBNull.Value)
                                                    {
                                                        y = Convert.ToString(dt_bend_welds.Rows[prev_i][col2]);
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
            dt_cc = null;
            dt_pmc = null;

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
                        if (W2 != null)
                        {
                            Wgen_main_form.dt_weld_map = Functions.Populate_data_table_from_excel(Wgen_main_form.dt_weld_map, W2, start_row, textBox_1.Text, textBox_2.Text, textBox_3.Text, textBox_4.Text, textBox_5.Text, textBox_6.Text, textBox_7.Text, textBox_8.Text, textBox_9.Text, textBox_10.Text, textBox_11.Text, true);
                            if (Wgen_main_form.dt_weld_map.Rows.Count > 0)
                            {
                                Wgen_main_form.tpage_weldmap.Hide();
                                Wgen_main_form.tpage_blank.Show();
                                Wgen_main_form.tpage_pipe_manifest.Hide();
                                Wgen_main_form.tpage_pipe_tally.Hide();
                                Wgen_main_form.tpage_allpts.Hide();
                                Wgen_main_form.tpage_build_pipe_tally.Hide();
                                Wgen_main_form.tpage_duplicates.Hide();
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

        private void Wgen_weldmap_Load(object sender, EventArgs e)
        {

        }
    }
}
