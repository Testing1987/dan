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
using Autodesk.AutoCAD.EditorInput;

namespace Alignment_mdi
{
    public partial class Agen_load_cl_from_xl : Form
    {

        System.Data.DataTable dt_errors;

        string Col_3DSta = "3DSta";
        string Col_BackSta = "BackSta";
        string Col_AheadSta = "AheadSta";

        string col_handle = "handle";
        string col_blockname = "Blockname";
        string col_x = "X";
        string col_y = "Y";
        string col_visibility = "Visibility";
        string col_dist1 = "Distance1";
        string col_dist2 = "Distance2";
        string col_dist3 = "Distance3";
        string col_dist4 = "Distance4";
        string col_sta = "Reference Station";

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();

            lista_butoane.Add(button_back_to_settings);
            lista_butoane.Add(button_calc_chainage_from_point);
            lista_butoane.Add(button_2D_to_3D);
            lista_butoane.Add(button_cl_l);
            lista_butoane.Add(button_cl_nl);
            lista_butoane.Add(button_create_centerline_file);
            lista_butoane.Add(button_draft_heavy_wall_based_on_chainage);
            lista_butoane.Add(button_generate_point_from_CSF);

            lista_butoane.Add(button_pick_points_output_chainages);
            lista_butoane.Add(button_point2sta);


            lista_butoane.Add(button_sta2point);


            lista_butoane.Add(button_verify_order);

            lista_butoane.Add(comboBox_ws1);
            lista_butoane.Add(textBox_chainage);
            lista_butoane.Add(textBox_col_csf);
            lista_butoane.Add(textBox_col_e);
            lista_butoane.Add(textBox_col_n);
            lista_butoane.Add(textBox_col_reroute);
            lista_butoane.Add(textBox_col_z);
            lista_butoane.Add(textBox_row_end);
            lista_butoane.Add(textBox_row_start);

            lista_butoane.Add(button_read_band_to_xl);
            lista_butoane.Add(button_calc_length);
            lista_butoane.Add(button_update_band);
            lista_butoane.Add(button_project_cl_bends_to_profile);
            lista_butoane.Add(button_arc_to_lines);
            lista_butoane.Add(button_P2C);




            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }
        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_back_to_settings);

            lista_butoane.Add(button_calc_chainage_from_point);
            lista_butoane.Add(button_2D_to_3D);
            lista_butoane.Add(button_cl_l);
            lista_butoane.Add(button_cl_nl);
            lista_butoane.Add(button_create_centerline_file);
            lista_butoane.Add(button_draft_heavy_wall_based_on_chainage);
            lista_butoane.Add(button_generate_point_from_CSF);

            lista_butoane.Add(button_pick_points_output_chainages);
            lista_butoane.Add(button_point2sta);
            lista_butoane.Add(button_sta2point);

            lista_butoane.Add(button_verify_order);

            lista_butoane.Add(comboBox_ws1);
            lista_butoane.Add(textBox_chainage);
            lista_butoane.Add(textBox_col_csf);
            lista_butoane.Add(textBox_col_e);
            lista_butoane.Add(textBox_col_n);
            lista_butoane.Add(textBox_col_reroute);
            lista_butoane.Add(textBox_col_z);
            lista_butoane.Add(textBox_row_end);
            lista_butoane.Add(textBox_row_start);

            lista_butoane.Add(button_read_band_to_xl);
            lista_butoane.Add(button_calc_length);
            lista_butoane.Add(button_update_band);
            lista_butoane.Add(button_project_cl_bends_to_profile);
            lista_butoane.Add(button_arc_to_lines);
            lista_butoane.Add(button_P2C);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }

        public Agen_load_cl_from_xl()
        {
            InitializeComponent();


        }

        private void TextBox_keypress_only_integers(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_integer_pozitive_at_keypress(sender, e);
        }

        private void append_deflections_to_dtcl(System.Data.DataTable dt1)
        {
            Plane Planul_curent = new Plane(new Point3d(0, 0, 0), Vector3d.ZAxis);
            if (dt1 != null && dt1.Rows.Count > 2 && dt1.Columns.Contains("X") == true && dt1.Columns.Contains("Y") == true && dt1.Columns.Contains("DeflAng") == true && dt1.Columns.Contains("DeflAngDMS") == true)
            {
                for (int i = 1; i < dt1.Rows.Count - 1; ++i)
                {
                    if (dt1.Rows[i]["X"] != DBNull.Value &&
                        dt1.Rows[i]["Y"] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["X"])) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["Y"])) == true &&
                        dt1.Rows[i - 1]["X"] != DBNull.Value &&
                        dt1.Rows[i - 1]["Y"] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i - 1]["X"])) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i - 1]["Y"])) == true &&
                        dt1.Rows[i + 1]["X"] != DBNull.Value &&
                        dt1.Rows[i + 1]["Y"] != DBNull.Value &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i + 1]["X"])) == true &&
                        Functions.IsNumeric(Convert.ToString(dt1.Rows[i + 1]["Y"])) == true)
                    {
                        double x1 = Convert.ToDouble(dt1.Rows[i - 1]["X"]);
                        double y1 = Convert.ToDouble(dt1.Rows[i - 1]["Y"]);
                        double x2 = Convert.ToDouble(dt1.Rows[i]["X"]);
                        double y2 = Convert.ToDouble(dt1.Rows[i]["Y"]);
                        double x3 = Convert.ToDouble(dt1.Rows[i + 1]["X"]);
                        double y3 = Convert.ToDouble(dt1.Rows[i + 1]["Y"]);
                        Vector3d vector1 = new Point3d(x1, y1, 0).GetVectorTo(new Point3d(x2, y2, 0));
                        Vector3d vector2 = new Point3d(x2, y2, 0).GetVectorTo(new Point3d(x3, y3, 0));
                        double Angle1 = (vector2.GetAngleTo(vector1)) * 180 / Math.PI;
                        string DMS1 = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);
                        dt1.Rows[i]["DeflAng"] = Angle1;
                        dt1.Rows[i]["DeflAngDMS"] = DMS1;
                    }
                }
            }
        }

        private void button_create_centerline_xl_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                return;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;


            if (System.IO.File.Exists(_AGEN_mainform.config_path) == true)
            {
                if (MessageBox.Show("all existing data will be overwriten... \r\nare you sure? ", "agen", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    return;
                }
            }
            else
            {
                MessageBox.Show("Please save your project before you load the centerline");
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    _AGEN_mainform.dt_centerline = Functions.Creaza_centerline_datatable_structure();

                    string Col_x = "X";
                    string Col_y = "Y";
                    string Col_z = "Z";
                    string Col_3DSta = "3DSta";
                    string Col_CSF = "CSF";
                    string Col_rr = "Reroute#";
                    _AGEN_mainform.dt_centerline.Columns.Add(Col_CSF, typeof(double));
                    _AGEN_mainform.dt_centerline.Columns.Add(Col_rr, typeof(string));


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
                                Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                if (W1 != null)
                                {
                                    _AGEN_mainform.dt_centerline = Functions.build_data_table_from_excel_based_on_columns_with_type_check(_AGEN_mainform.dt_centerline, W1, start1, end1, Col_x, textBox_col_e.Text,
                                        Col_y, textBox_col_n.Text, Col_z, textBox_col_z.Text, Col_3DSta, textBox_chainage.Text, Col_CSF, textBox_col_csf.Text, Col_rr, textBox_col_reroute.Text);

                                    append_deflections_to_dtcl(_AGEN_mainform.dt_centerline);

                                    dt_errors = new System.Data.DataTable();
                                    dt_errors.Columns.Add("Survey File Value", typeof(double));
                                    dt_errors.Columns.Add("Calculated Value", typeof(double));
                                    dt_errors.Columns.Add("Error Type", typeof(string));

                                    _AGEN_mainform.version = filename;
                                    Functions.create_backup(fisier_cl);

                                    _AGEN_mainform.tpage_setup.Populate_centerline_file(fisier_cl, true, true);
                                    _AGEN_mainform.tpage_setup.Set_centerline_label_to_green();

                                    button_cl_l.Visible = true;
                                    button_cl_nl.Visible = false;

                                    double sta_cumul = 0;

                                    if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                    {
                                        for (int i = 1; i < _AGEN_mainform.dt_centerline.Rows.Count; ++i)
                                        {
                                            double X1 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i - 1]["X"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1]["X"])) == true)
                                            {
                                                X1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1]["X"]);
                                            }
                                            double Y1 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i - 1]["Y"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1]["Y"])) == true)
                                            {
                                                Y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1]["Y"]);
                                            }
                                            double Z1 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i - 1]["Z"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1]["Z"])) == true)
                                            {
                                                Z1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1]["Z"]);
                                            }
                                            double CSF1 = 1;
                                            if (_AGEN_mainform.dt_centerline.Rows[i - 1]["CSF"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1]["CSF"])) == true)
                                            {
                                                CSF1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1]["CSF"]);
                                            }
                                            double STA3D1 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i - 1]["3DSta"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i - 1]["3DSta"])) == true)
                                            {
                                                STA3D1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i - 1]["3DSta"]);
                                            }
                                            if (i == 1) sta_cumul = STA3D1;
                                            double X2 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i]["X"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i]["X"])) == true)
                                            {
                                                X2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i]["X"]);
                                            }
                                            double Y2 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i]["Y"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i]["Y"])) == true)
                                            {
                                                Y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i]["Y"]);
                                            }
                                            double Z2 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i]["Z"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i]["Z"])) == true)
                                            {
                                                Z2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i]["Z"]);
                                            }
                                            double CSF2 = 1;
                                            if (_AGEN_mainform.dt_centerline.Rows[i]["CSF"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i]["CSF"])) == true)
                                            {
                                                CSF2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i]["CSF"]);
                                            }
                                            double STA3D2 = -1.234;
                                            if (_AGEN_mainform.dt_centerline.Rows[i]["3DSta"] != DBNull.Value &&
                                                Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[i]["3DSta"])) == true)
                                            {
                                                STA3D2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[i]["3DSta"]);
                                            }

                                            double calc_sta = -1.234;

                                            if (X1 != -1.234 && X2 != -1.234 && Y1 != -1.234 && Y2 != -1.234 && Z1 != -1.234 && Z2 != -1.234)
                                            {
                                                calc_sta = STA3D1 + Math.Pow(Math.Pow(X1 - X2, 2) + Math.Pow(Y1 - Y2, 2) + Math.Pow(Z1 - Z2, 2), 0.5) / ((CSF1 + CSF2) / 2);
                                                sta_cumul = sta_cumul + Math.Pow(Math.Pow(X1 - X2, 2) + Math.Pow(Y1 - Y2, 2) + Math.Pow(Z1 - Z2, 2), 0.5) / ((CSF1 + CSF2) / 2);

                                            }
                                            else
                                            {
                                                double VAL1 = STA3D1;
                                                if (VAL1 == -1.234) VAL1 = STA3D2;

                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Survey File Value"] = VAL1;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Gap in data";
                                            }


                                            double amt = 0.1;

                                            if (_AGEN_mainform.round1 == 1) amt = 0.01;
                                            if (_AGEN_mainform.round1 == 2) amt = 0.001;
                                            if (_AGEN_mainform.round1 == 3) amt = 0.0001;

                                            if (Math.Abs(calc_sta - STA3D2) > amt)
                                            {
                                                dt_errors.Rows.Add();
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Survey File Value"] = STA3D2;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = calc_sta;
                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Point to Point distance";

                                            }
                                            if (i == _AGEN_mainform.dt_centerline.Rows.Count - 1)
                                            {
                                                if (Math.Abs(sta_cumul - STA3D2) > amt)
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Survey File Value"] = STA3D2;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = sta_cumul;
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Total calc distance";
                                                }
                                            }
                                        }
                                        transfer_errors_to_panel(dt_errors);

                                        _AGEN_mainform.COUNTRY = "CANADA";
                                        _AGEN_mainform.tpage_setup.set_radioButton_canada(true);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void make_first_line_invisible()
        {


            dataGridView_errors.DataSource = null;
        }

        private void transfer_errors_to_panel(System.Data.DataTable dt1)
        {
            if (dt1.Rows.Count > 0)
            {
                dataGridView_errors.DataSource = dt1;
                dataGridView_errors.Columns[0].Width = 150;
                dataGridView_errors.Columns[1].Width = 150;
                dataGridView_errors.Columns[2].Width = 300;
                dataGridView_errors.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_errors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView_errors.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                dataGridView_errors.DefaultCellStyle.ForeColor = Color.White;
                dataGridView_errors.EnableHeadersVisualStyles = false;
            }
        }



        private void button_back_to_settings_Click(object sender, EventArgs e)
        {
            _AGEN_mainform.tpage_processing.Hide();
            _AGEN_mainform.tpage_blank.Hide();
            _AGEN_mainform.tpage_setup.Show();
            _AGEN_mainform.tpage_viewport_settings.Hide();
            _AGEN_mainform.tpage_tblk_attrib.Hide();
            _AGEN_mainform.tpage_band_analize.Hide();
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
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();

            _AGEN_mainform.tpage_tools.Hide();
            _AGEN_mainform.tpage_st_eq.Hide();
            _AGEN_mainform.tpage_cl_xl.Hide();
        }

        private void button_export_errors_to_xl_Click(object sender, EventArgs e)
        {
            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt_errors);
        }

        private void button_2D_3D_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        ObjectId[] Empty_array = null;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        set_enable_false();

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                                Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                Polyline3d poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_z = "Z";
                                string Col_Sta2D = "2D distance";
                                string Col_Sta3D = "3D Chainage";

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
                                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                            if (W1 != null)
                                            {
                                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                                dt1.Columns.Add(Col_x, typeof(double));
                                                dt1.Columns.Add(Col_y, typeof(double));
                                                dt1.Columns.Add(Col_z, typeof(double));
                                                dt1.Columns.Add(Col_Sta2D, typeof(double));
                                                dt1.Columns.Add(Col_Sta3D, typeof(double));
                                                dt1.Columns.Add("2Dpoint", typeof(string));

                                                List<string> lista_col = new List<string>();
                                                List<string> lista_colxl = new List<string>();
                                                lista_col.Add(Col_Sta2D);

                                                lista_colxl.Add(textBox_chainage.Text);


                                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, lista_col, lista_colxl);
                                                dt_errors = new System.Data.DataTable();
                                                dt_errors.Columns.Add("Survey File Value", typeof(double));
                                                dt_errors.Columns.Add("Calculated Value", typeof(string));
                                                dt_errors.Columns.Add("Error Type", typeof(string));

                                                if (dt1.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        double dist2d = -1;

                                                        if (dt1.Rows[i][Col_Sta2D] != DBNull.Value)
                                                        {
                                                            dist2d = Convert.ToDouble(dt1.Rows[i][Col_Sta2D]);
                                                        }

                                                        if (dist2d >= 0 && dist2d <= poly2D.Length)
                                                        {
                                                            Point3d pt2d = poly2D.GetPointAtDist(dist2d);
                                                            double x = pt2d.X;
                                                            double y = pt2d.Y;
                                                            double b1 = -1.23456;
                                                            double sta = Functions.get_stationCSF_from_point(poly2D, pt2d, dist2d, _AGEN_mainform.dt_centerline, ref b1);

                                                            dt1.Rows[i][Col_Sta2D] = dist2d;
                                                            dt1.Rows[i][Col_Sta3D] = sta;

                                                            dt1.Rows[i]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                            Point3d pt_on_poly = poly2D.GetClosestPointTo(new Point3d(x, y, poly2D.Elevation), Vector3d.ZAxis, false);
                                                            double param1 = poly2D.GetParameterAtPoint(pt_on_poly);
                                                            if (param1 > poly3D.EndParam) param1 = poly3D.EndParam;
                                                            Point3d point_for_z = poly3D.GetPointAtParameter(param1);
                                                            dt1.Rows[i][Col_z] = point_for_z.Z;
                                                            dt1.Rows[i][Col_x] = x;
                                                            dt1.Rows[i][Col_y] = y;
                                                            if (b1 != -1.23456)
                                                            {
                                                                if (dt1.Columns.Contains("BacK Station") == false) dt1.Columns.Add("BacK Station", typeof(double));
                                                                if (dt1.Columns.Contains("Ahead Station") == false) dt1.Columns.Add("Ahead Station", typeof(double));
                                                                dt1.Rows[i]["Ahead Station"] = sta;
                                                                dt1.Rows[i]["BacK Station"] = b1;
                                                            }

                                                        }

                                                        else
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not valid station";
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = "See Row " + (start1 + i).ToString();
                                                        }
                                                    }
                                                    transfer_errors_to_panel(dt_errors);
                                                    dt1.Columns.Add("Segment", typeof(string));
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        dt1.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                    }
                                                    string nume1 = System.DateTime.Now.Hour + "-" + System.DateTime.Now.Minute;
                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1);
                                                }
                                            }
                                        }
                                    }
                                }
                                poly3D.Erase();
                                Trans1.Commit();
                            }
                        }
                        Editor1.SetImpliedSelection(Empty_array);
                        Editor1.WriteMessage("\nCommand:");
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_generate_point_from_CSF_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                        Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                        string Col_x = "X";
                        string Col_y = "Y";
                        string Col_z = "Z";
                        string Col_Sta = "Chainage";

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
                                    Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                    if (W1 != null)
                                    {
                                        System.Data.DataTable dt1 = new System.Data.DataTable();
                                        dt1.Columns.Add(Col_x, typeof(double));
                                        dt1.Columns.Add(Col_y, typeof(double));
                                        dt1.Columns.Add(Col_z, typeof(double));
                                        dt1.Columns.Add(Col_Sta, typeof(string));
                                        dt1.Columns.Add("2Dpoint", typeof(string));


                                        System.Data.DataRow[] rows1 = null;
                                        int number_of_rows = 0;
                                        List<string> lista_col = new List<string>();
                                        List<string> lista_colxl = new List<string>();

                                        lista_col.Add(Col_Sta);

                                        lista_colxl.Add(textBox_chainage.Text);
                                        dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, lista_col, lista_colxl);

                                        dt_errors = new System.Data.DataTable();
                                        dt_errors.Columns.Add("Survey File Value", typeof(double));
                                        dt_errors.Columns.Add("Calculated Value", typeof(double));
                                        dt_errors.Columns.Add("Error Type", typeof(string));


                                        if (dt1.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {

                                                if (dt1.Rows[i][Col_Sta] != DBNull.Value &&
                                                    Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_Sta]).Replace("+", "")) == true)
                                                {
                                                    double sta = Convert.ToDouble(Convert.ToString(dt1.Rows[i][Col_Sta]).Replace("+", ""));
                                                    if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                                    {
                                                        for (int j = 0; j < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++j)
                                                        {
                                                            if (_AGEN_mainform.dt_centerline.Rows[j]["3DSta"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                                _AGEN_mainform.dt_centerline.Rows[j]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j]["Y"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j]["Z"] != DBNull.Value &&
                                                                _AGEN_mainform.dt_centerline.Rows[j + 1]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["Y"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["Z"] != DBNull.Value)
                                                            {
                                                                double sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j][Col_3DSta]);
                                                                double sta2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1][Col_3DSta]);

                                                                if (_AGEN_mainform.dt_centerline.Rows[j][Col_AheadSta] != DBNull.Value)
                                                                {
                                                                    sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j][Col_AheadSta]);
                                                                }


                                                                if (_AGEN_mainform.dt_centerline.Rows[j + 1][Col_BackSta] != DBNull.Value)
                                                                {
                                                                    sta2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1][Col_BackSta]);
                                                                }

                                                                if (sta >= sta1 && sta <= sta2)
                                                                {


                                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["X"]);
                                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Y"]);
                                                                    double z1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Z"]);
                                                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"]);
                                                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"]);
                                                                    double z2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Z"]);

                                                                    double x = x1 + (x2 - x1) * (sta - sta1) / (sta2 - sta1);
                                                                    double y = y1 + (y2 - y1) * (sta - sta1) / (sta2 - sta1);
                                                                    double z = z1 + (z2 - z1) * (sta - sta1) / (sta2 - sta1);


                                                                    if (dt1.Rows[i][Col_x] == DBNull.Value)
                                                                    {
                                                                        dt1.Rows[i][Col_x] = x;
                                                                        dt1.Rows[i][Col_y] = y;
                                                                        dt1.Rows[i]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                                        dt1.Rows[i][Col_z] = z;
                                                                        dt1.Rows[i][Col_Sta] = Functions.Get_chainage_from_double(sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                                    }
                                                                    else
                                                                    {
                                                                        ++number_of_rows;
                                                                        Array.Resize(ref rows1, number_of_rows);
                                                                        rows1[number_of_rows - 1] = dt1.NewRow();

                                                                        rows1[number_of_rows - 1][Col_x] = x;
                                                                        rows1[number_of_rows - 1][Col_y] = y;
                                                                        rows1[number_of_rows - 1][Col_z] = z;
                                                                        rows1[number_of_rows - 1][Col_Sta] = Functions.Get_chainage_from_double(sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                                                        rows1[number_of_rows - 1]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                dt_errors.Rows.Add();
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric values omn centerline.xls on row " + (j + _AGEN_mainform.Start_row_CL).ToString();
                                                            }
                                                        }



                                                    }
                                                }
                                                else
                                                {
                                                    dt_errors.Rows.Add();
                                                    dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric chainage on row " + (i + start1).ToString();
                                                }

                                            }


                                            if (number_of_rows > 0)
                                            {

                                                for (int i = 0; i < rows1.Length; ++i)
                                                {
                                                    dt1.Rows.Add(rows1[i]);
                                                }

                                                dt1.Columns.Add("Duplicate Point", typeof(string));

                                                for (int i = dt1.Rows.Count - 1; i >= number_of_rows; --i)
                                                {
                                                    dt1.Rows[i]["Duplicate Point"] = "YES";
                                                }
                                            }


                                            transfer_errors_to_panel(dt_errors);
                                            dt1.Columns.Add("Segment", typeof(string));
                                            string label1 = _AGEN_mainform.current_segment;
                                            if (label1 == "")
                                            {
                                                label1 = _AGEN_mainform.tpage_setup.get_textBox_client_name_content() + " - " + _AGEN_mainform.tpage_setup.get_textBox_project_name_content();
                                            }

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                dt1.Rows[i]["Segment"] = label1;
                                            }
                                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }
        private void button_calc_chainage_from_point_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        ObjectId[] Empty_array = null;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        set_enable_false();

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                                Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                Polyline3d poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_z = "Z";
                                string Col_Sta = "Calculated Rounded Chainage";
                                string Col_sta_raw = "RAW Chainage";
                                string Col_offset = "2D Offset";

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
                                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                            if (W1 != null)
                                            {
                                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                                dt1.Columns.Add(Col_x, typeof(double));
                                                dt1.Columns.Add(Col_y, typeof(double));
                                                dt1.Columns.Add(Col_z, typeof(double));
                                                dt1.Columns.Add(Col_Sta, typeof(double));
                                                dt1.Columns.Add(Col_sta_raw, typeof(double));
                                                dt1.Columns.Add(Col_offset, typeof(double));
                                                dt1.Columns.Add("2Dpoint", typeof(string));

                                                List<string> lista_col = new List<string>();
                                                List<string> lista_colxl = new List<string>();
                                                lista_col.Add(Col_x);
                                                lista_col.Add(Col_y);
                                                lista_colxl.Add(textBox_col_e.Text);
                                                lista_colxl.Add(textBox_col_n.Text);

                                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, lista_col, lista_colxl);
                                                dt_errors = new System.Data.DataTable();
                                                dt_errors.Columns.Add("Survey File Value", typeof(double));
                                                dt_errors.Columns.Add("Calculated Value", typeof(string));
                                                dt_errors.Columns.Add("Error Type", typeof(string));

                                                if (dt1.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        double x = -1.234;
                                                        if (dt1.Rows[i][Col_x] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_x])) == true)
                                                        {
                                                            x = Convert.ToDouble(dt1.Rows[i][Col_x]);
                                                        }
                                                        double y = -1.234;
                                                        if (dt1.Rows[i][Col_y] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_y])) == true)
                                                        {
                                                            y = Convert.ToDouble(dt1.Rows[i][Col_y]);
                                                        }
                                                        Point3d pt2d = poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                                        double dist2d = poly2D.GetDistAtPoint(pt2d);
                                                        if (x != -1.234 && y != -1.234)
                                                        {
                                                            double b1 = -1.23456;
                                                            double sta = Functions.get_stationCSF_from_point(poly2D, pt2d, dist2d, _AGEN_mainform.dt_centerline, ref b1);
                                                            double calc_sta = Math.Round(sta, _AGEN_mainform.round1);
                                                            dt1.Rows[i][Col_Sta] = calc_sta;
                                                            dt1.Rows[i][Col_sta_raw] = sta;
                                                            dt1.Rows[i]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                            Point3d pt_on_poly = poly2D.GetClosestPointTo(new Point3d(x, y, poly2D.Elevation), Vector3d.ZAxis, false);
                                                            double param1 = poly2D.GetParameterAtPoint(pt_on_poly);
                                                            if (param1 > poly3D.EndParam) param1 = poly3D.EndParam;
                                                            Point3d point_for_z = poly3D.GetPointAtParameter(param1);
                                                            dt1.Rows[i][Col_z] = point_for_z.Z;
                                                            double offset1 = Math.Pow(Math.Pow(x - pt_on_poly.X, 2) + Math.Pow(y - pt_on_poly.Y, 2), 0.5);
                                                            dt1.Rows[i][Col_offset] = offset1;
                                                            if (b1 != -1.23456)
                                                            {
                                                                if (dt1.Columns.Contains("BacK Station") == false) dt1.Columns.Add("BacK Station", typeof(double));
                                                                if (dt1.Columns.Contains("Ahead Station") == false) dt1.Columns.Add("Ahead Station", typeof(double));
                                                                dt1.Rows[i]["Ahead Station"] = sta;
                                                                dt1.Rows[i]["BacK Station"] = b1;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric x or y";
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = "See Row " + (start1 + i).ToString();
                                                        }
                                                    }
                                                    transfer_errors_to_panel(dt_errors);
                                                    dt1.Columns.Add("Segment", typeof(string));
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        dt1.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                    }
                                                    string nume1 = System.DateTime.Now.Hour + "-" + System.DateTime.Now.Minute;
                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1);
                                                }
                                            }
                                        }
                                    }
                                }
                                poly3D.Erase();
                                Trans1.Commit();
                            }
                        }
                        Editor1.SetImpliedSelection(Empty_array);
                        Editor1.WriteMessage("\nCommand:");
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_pick_points_output_chainages_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();


            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }

                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not exist");

                    return;
                }

                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not have any data");
                    return;
                }


            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }




            System.Data.DataTable Dt_poly = Functions.Creaza_prof_poly_dt_structure();




            try
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("STA1", typeof(double));
                        dt1.Columns.Add("STA2", typeof(double));
                        dt1.Columns.Add("X1", typeof(double));
                        dt1.Columns.Add("Y1", typeof(double));
                        dt1.Columns.Add("X2", typeof(double));
                        dt1.Columns.Add("Y2", typeof(double));
                        dataGridView_errors.DataSource = dt1;
                        dataGridView_errors.Columns[0].Width = 75;
                        dataGridView_errors.Columns[1].Width = 75;
                        dataGridView_errors.Columns[2].Width = 125;
                        dataGridView_errors.Columns[3].Width = 125;
                        dataGridView_errors.Columns[4].Width = 125;
                        dataGridView_errors.Columns[5].Width = 125;
                        dataGridView_errors.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_errors.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                        dataGridView_errors.DefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
                        dataGridView_errors.DefaultCellStyle.ForeColor = Color.White;
                        dataGridView_errors.EnableHeadersVisualStyles = false;

                        _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                        bool next1 = true;
                        bool pick_first = true;
                        Point3d prev_pt = new Point3d();
                        do
                        {



                            Point3d pt1 = new Point3d();

                            if (pick_first == true)
                            {

                                Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify start point");
                                PP1.AllowNone = false;

                                Point_res1 = Editor1.GetPoint(PP1);

                                if (Point_res1.Status != PromptStatus.OK)
                                {

                                    Trans1.Commit();
                                    dt1 = Functions.Sort_data_table(dt1, "STA1");
                                    Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);

                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    dataGridView_errors.DataSource = null;
                                    set_enable_true();
                                    return;
                                }

                                pt1 = Point_res1.Value;
                            }
                            else
                            {
                                pt1 = prev_pt;
                            }

                            if (checkBox_chain.Checked == true)
                            {
                                pick_first = false;
                            }
                            else
                            {
                                pick_first = true;
                            }


                            Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res2;
                            Autodesk.AutoCAD.EditorInput.PromptPointOptions PP2;
                            PP2 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nSpecify end point");
                            PP2.AllowNone = false;
                            PP2.UseBasePoint = true;
                            PP2.BasePoint = pt1;
                            Point_res2 = Editor1.GetPoint(PP2);

                            if (Point_res2.Status != PromptStatus.OK)
                            {

                                Trans1.Commit();
                                dt1 = Functions.Sort_data_table(dt1, "STA1");
                                Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);

                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                dataGridView_errors.DataSource = null;
                                set_enable_true();
                                return;
                            }

                            Point3d pt2 = Point_res2.Value;
                            prev_pt = pt2;

                            Point3d point_on_poly1 = _AGEN_mainform.Poly2D.GetClosestPointTo(pt1, Vector3d.ZAxis, false);
                            Point3d point_on_poly2 = _AGEN_mainform.Poly2D.GetClosestPointTo(pt2, Vector3d.ZAxis, false);
                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(point_on_poly1);
                            double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(point_on_poly2);

                            if (param1 > param2)
                            {
                                Point3d t = point_on_poly1;
                                point_on_poly1 = point_on_poly2;
                                point_on_poly2 = t;
                                double tt = param1;
                                param1 = param2;
                                param2 = tt;

                            }

                            dt1.Rows.Add();
                            dt1.Rows[dt1.Rows.Count - 1]["X1"] = point_on_poly1.X;
                            dt1.Rows[dt1.Rows.Count - 1]["Y1"] = point_on_poly1.Y;
                            dt1.Rows[dt1.Rows.Count - 1]["X2"] = point_on_poly2.X;
                            dt1.Rows[dt1.Rows.Count - 1]["Y2"] = point_on_poly2.Y;



                            if (_AGEN_mainform.COUNTRY == "CANADA")
                            {
                                double b11 = -1.23456;
                                double b21 = -1.23456;
                                double d1_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                double d2_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param2);
                                double Sta1 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, point_on_poly1, d1_2d, _AGEN_mainform.dt_centerline, ref b11);
                                double Sta2 = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, point_on_poly2, d2_2d, _AGEN_mainform.dt_centerline, ref b21);
                                dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(Sta1, _AGEN_mainform.round1);
                                dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(Sta2, _AGEN_mainform.round1);
                            }
                            else

                            {
                                double d1 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                double d2 = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param2);

                                dt1.Rows[dt1.Rows.Count - 1]["STA1"] = Math.Round(d1, _AGEN_mainform.round1);
                                dt1.Rows[dt1.Rows.Count - 1]["STA2"] = Math.Round(d2, _AGEN_mainform.round1);
                            }
                        } while (next1 == true);


                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            set_enable_true();


        }

        private void button_point2sta_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        try
                        {
                            set_enable_false();
                            using (DocumentLock lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                    _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                    string Col_x = "X";
                                    string Col_y = "Y";
                                    string Col_Sta = "Measured Station";
                                    string Col_steq_Sta = "Equated Station";
                                    string Col_offset = "2D Offset";
                                    string Col_Sta_raw = "Measured Station Raw";

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
                                                Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                                if (W1 != null)
                                                {
                                                    System.Data.DataTable dt1 = new System.Data.DataTable();
                                                    dt1.Columns.Add(Col_x, typeof(string));
                                                    dt1.Columns.Add(Col_y, typeof(string));
                                                    dt1.Columns.Add(Col_Sta, typeof(string));
                                                    dt1.Columns.Add(Col_steq_Sta, typeof(string));
                                                    dt1.Columns.Add(Col_offset, typeof(string));
                                                    dt1.Columns.Add(Col_Sta_raw, typeof(string));


                                                    dt1 = Functions.build_data_table_from_excel_based_on_columns_with_type_check(dt1, W1,
                                                                    start1, end1,
                                                                    Col_x, textBox_col_e.Text,
                                                                    Col_y, textBox_col_n.Text,
                                                                    textBox_chainage.Text, "",
                                                                    "", "", "", "", "", "");

                                                    dt_errors = new System.Data.DataTable();
                                                    dt_errors.Columns.Add("Survey File Value", typeof(double));
                                                    dt_errors.Columns.Add("Calculated Value", typeof(string));
                                                    dt_errors.Columns.Add("Error Type", typeof(string));

                                                    if (dt1.Rows.Count > 0)
                                                    {
                                                        for (int i = 0; i < dt1.Rows.Count; ++i)
                                                        {
                                                            double x = -1.234;
                                                            if (dt1.Rows[i][Col_x] != DBNull.Value &&
                                                                Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_x])) == true)
                                                            {
                                                                x = Convert.ToDouble(dt1.Rows[i][Col_x]);
                                                            }

                                                            double y = -1.234;
                                                            if (dt1.Rows[i][Col_y] != DBNull.Value &&
                                                                Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_y])) == true)
                                                            {
                                                                y = Convert.ToDouble(dt1.Rows[i][Col_y]);
                                                            }

                                                            Point3d pt2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt2d);
                                                            if (param1 > _AGEN_mainform.Poly3D.EndParam) param1 = _AGEN_mainform.Poly3D.EndParam;

                                                            if (x != -1.234 && y != -1.234)
                                                            {
                                                                double calc_sta = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                                                dt1.Rows[i][Col_Sta] = Convert.ToString(Math.Round(calc_sta, _AGEN_mainform.round1));
                                                                dt1.Rows[i][Col_Sta_raw] = Convert.ToString(calc_sta);
                                                                if (_AGEN_mainform.dt_station_equation != null)
                                                                {
                                                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                                    {
                                                                        double steq = Functions.Station_equation_ofV2(Math.Round(calc_sta, 3), _AGEN_mainform.dt_station_equation);
                                                                        dt1.Rows[i][Col_steq_Sta] = Convert.ToString(steq);
                                                                    }
                                                                }

                                                                Point3d pt_on_poly = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                                                double offset1 = Math.Pow(Math.Pow(x - pt_on_poly.X, 2) + Math.Pow(y - pt_on_poly.Y, 2), 0.5);
                                                                dt1.Rows[i][Col_offset] = Convert.ToString(offset1);

                                                            }
                                                            else
                                                            {
                                                                dt_errors.Rows.Add();
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric x or y";
                                                                dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = "See Row " + (start1 + i).ToString();
                                                            }
                                                        }

                                                        transfer_errors_to_panel(dt_errors);
                                                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    _AGEN_mainform.Poly3D.Erase();

                                    Trans1.Commit();
                                }
                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }


                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_sta2point_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                        Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                        _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        try
                        {
                            set_enable_false();
                            using (DocumentLock lock1 = ThisDrawing.LockDocument())
                            {
                                using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                                {
                                    BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                    BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                    _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                    _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                    string Col_x = "X";
                                    string Col_y = "Y";
                                    string Col_z = "Z";
                                    string Col_Sta = "Station";
                                    string Col_steq_Sta = "Equated Station";


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
                                                Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                                if (W1 != null)
                                                {
                                                    System.Data.DataTable dt1 = new System.Data.DataTable();
                                                    dt1.Columns.Add(Col_x, typeof(string));
                                                    dt1.Columns.Add(Col_y, typeof(string));
                                                    dt1.Columns.Add(Col_z, typeof(string));
                                                    dt1.Columns.Add(Col_Sta, typeof(string));
                                                    dt1.Columns.Add(Col_steq_Sta, typeof(string));




                                                    dt1 = Functions.build_data_table_from_excel_based_on_columns_with_type_check(dt1, W1,
                                                                    start1, end1,
                                                                    "", "", "", "", "", "",
                                                                    Col_Sta, textBox_chainage.Text,
                                                                    "", "", "", "");


                                                    if (dt1.Rows.Count > 0)
                                                    {
                                                        int i = 0;
                                                        int xtra = 0;

                                                        do
                                                        {
                                                            if (dt1.Rows[i][Col_Sta] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_Sta]).Replace("+", "")) == true)
                                                            {
                                                                double sta = Convert.ToDouble(Convert.ToString(dt1.Rows[i][Col_Sta]).Replace("+", ""));
                                                                xtra = 0;

                                                                if (_AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                                {
                                                                    List<double> lista_measured = Functions.Equation_to_measured(sta, _AGEN_mainform.Poly3D, _AGEN_mainform.Poly2D, _AGEN_mainform.dt_station_equation);
                                                                    dt1.Rows[i][Col_steq_Sta] = sta;
                                                                    dt1.Rows[i][Col_Sta] = Convert.ToString(lista_measured[0]);

                                                                    if (lista_measured.Count > 1)
                                                                    {
                                                                        for (int j = 1; j < lista_measured.Count; ++j)
                                                                        {
                                                                            ++xtra;
                                                                            System.Data.DataRow row1 = dt1.NewRow();
                                                                            for (int k = 0; k < dt1.Columns.Count; ++k)
                                                                            {
                                                                                row1[k] = dt1.Rows[i][k];
                                                                            }
                                                                            row1[Col_Sta] = Convert.ToString(lista_measured[j]);
                                                                            row1[Col_steq_Sta] = sta;
                                                                            dt1.Rows.InsertAt(row1, i + 1);
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    dt1.Rows[i][Col_Sta] = sta;
                                                                }
                                                            }


                                                            i = i + 1 + xtra;
                                                        } while (i < dt1.Rows.Count);
                                                    }

                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        double sta1 = Convert.ToDouble(dt1.Rows[i][Col_Sta]);
                                                        if (sta1 >= 0 && sta1 < +_AGEN_mainform.Poly3D.Length)
                                                        {
                                                            dt1.Rows[i][Col_x] = _AGEN_mainform.Poly3D.GetPointAtDist(sta1).X;
                                                            dt1.Rows[i][Col_y] = _AGEN_mainform.Poly3D.GetPointAtDist(sta1).Y;
                                                            dt1.Rows[i][Col_z] = _AGEN_mainform.Poly3D.GetPointAtDist(sta1).Z;
                                                        }


                                                    }

                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
                                                }
                                            }
                                        }
                                    }
                                    _AGEN_mainform.Poly3D.Erase();

                                    Trans1.Commit();
                                }

                            }
                        }


                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_draft_heavy_wall_based_on_chainage_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            this.MdiParent.WindowState = FormWindowState.Minimized;
            set_enable_false();

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {
                string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
                {
                    ProjF = ProjF + "\\";
                }
                string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

                if (System.IO.File.Exists(fisier_cl) == false)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not exist");
                    return;
                }
                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    set_enable_true();
                    MessageBox.Show("the centerline file does not have any data");
                    return;
                }
            }
            else
            {
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }






            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }


            try
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                ObjectId[] Empty_array = null;
                Editor1.SetImpliedSelection(Empty_array);

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite);
                        LayerTable layertable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add("sta1", typeof(string));
                        dt1.Columns.Add("sta2", typeof(string));
                        dt1.Columns.Add("x1", typeof(double));
                        dt1.Columns.Add("y1", typeof(double));
                        dt1.Columns.Add("x2", typeof(double));
                        dt1.Columns.Add("y2", typeof(double));
                        dt1.Columns.Add("layer", typeof(string));

                        short ci = 1;


                        _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

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
                                    Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                    if (W1 != null)
                                    {


                                        dt1 = Functions.build_data_table_from_excel_based_on_columns_with_type_check(dt1, W1,
                                                        start1, end1,
                                                        "sta1", textBoxH1.Text, "sta2", textBoxH2.Text,
                                                        "layer", textBoxH3.Text, "", "", "", "", "", "");

                                        #region CANADA
                                        if (dt1.Rows.Count > 0 && _AGEN_mainform.COUNTRY == "CANADA")
                                        {
                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {

                                                if (dt1.Rows[i]["sta1"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["sta1"]).Replace("+", "")) == true &&
                                                    dt1.Rows[i]["sta2"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["sta2"]).Replace("+", "")) == true)
                                                {
                                                    double staH1 = Convert.ToDouble(Convert.ToString(dt1.Rows[i]["sta1"]).Replace("+", ""));
                                                    double staH2 = Convert.ToDouble(Convert.ToString(dt1.Rows[i]["sta2"]).Replace("+", ""));

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
                                                                double sta1 = Convert.ToDouble(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j]["3DSta"]).Replace("+", ""));
                                                                double sta2 = Convert.ToDouble(Convert.ToString(_AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"]).Replace("+", ""));
                                                                if (staH1 >= sta1 && staH1 <= sta2)
                                                                {
                                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["X"]);
                                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Y"]);

                                                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"]);
                                                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"]);

                                                                    double x = x1 + (x2 - x1) * (staH1 - sta1) / (sta2 - sta1);
                                                                    double y = y1 + (y2 - y1) * (staH1 - sta1) / (sta2 - sta1);

                                                                    dt1.Rows[i]["x1"] = x;
                                                                    dt1.Rows[i]["y1"] = y;
                                                                }

                                                                if (staH2 >= sta1 && staH2 <= sta2)
                                                                {
                                                                    double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["X"]);
                                                                    double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Y"]);
                                                                    double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"]);
                                                                    double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"]);
                                                                    double x = x1 + (x2 - x1) * (staH2 - sta1) / (sta2 - sta1);
                                                                    double y = y1 + (y2 - y1) * (staH2 - sta1) / (sta2 - sta1);
                                                                    dt1.Rows[i]["x2"] = x;
                                                                    dt1.Rows[i]["y2"] = y;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                double x1 = Convert.ToDouble(dt1.Rows[i]["x1"]);
                                                double y1 = Convert.ToDouble(dt1.Rows[i]["y1"]);
                                                double x2 = Convert.ToDouble(dt1.Rows[i]["x2"]);
                                                double y2 = Convert.ToDouble(dt1.Rows[i]["y2"]);
                                                Point3d pt1 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                                Point3d pt2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);

                                                double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt1);
                                                double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt2);
                                                Polyline poly_heavy_wall = Functions.get_part_of_poly(_AGEN_mainform.Poly2D, param1, param2);

                                                string layer1 = "0";
                                                if (dt1.Rows[i]["layer"] != DBNull.Value)
                                                {
                                                    layer1 = Convert.ToString(dt1.Rows[i]["layer"]);
                                                }


                                                if (layertable1.Has(layer1) == false)
                                                {
                                                    Functions.Creaza_layer(layer1, ci, true);
                                                    ++ci;
                                                    if (ci == 7) ci = 1;
                                                }
                                                poly_heavy_wall.Layer = layer1;

                                                BTrecord.AppendEntity(poly_heavy_wall);
                                                Trans1.AddNewlyCreatedDBObject(poly_heavy_wall, true);
                                            }
                                        }
                                        #endregion

                                        #region USA
                                        if (dt1.Rows.Count > 0 && _AGEN_mainform.COUNTRY == "USA")
                                        {
                                            _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {

                                                if (dt1.Rows[i]["sta1"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["sta1"]).Replace("+", "")) == true &&
                                                    dt1.Rows[i]["sta2"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[i]["sta2"]).Replace("+", "")) == true)
                                                {
                                                    double sta1 = Convert.ToDouble(Convert.ToString(dt1.Rows[i]["sta1"]).Replace("+", ""));
                                                    double sta2 = Convert.ToDouble(Convert.ToString(dt1.Rows[i]["sta2"]).Replace("+", ""));

                                                    if (sta1 >= _AGEN_mainform.Poly3D.Length)
                                                    {
                                                        sta1 = _AGEN_mainform.Poly3D.Length - 0.00001;
                                                    }

                                                    if (sta2 >= _AGEN_mainform.Poly3D.Length)
                                                    {
                                                        sta2 = _AGEN_mainform.Poly3D.Length - 0.00001;
                                                    }

                                                    if (sta1 >= 0 && sta1 <= _AGEN_mainform.Poly3D.Length)
                                                    {
                                                        Point3d pt1 = _AGEN_mainform.Poly3D.GetPointAtDist(sta1);
                                                        dt1.Rows[i]["x1"] = pt1.X;
                                                        dt1.Rows[i]["y1"] = pt1.Y;

                                                    }
                                                    if (sta2 >= 0 && sta2 <= _AGEN_mainform.Poly3D.Length)
                                                    {
                                                        Point3d pt2 = _AGEN_mainform.Poly3D.GetPointAtDist(sta2);
                                                        dt1.Rows[i]["x2"] = pt2.X;
                                                        dt1.Rows[i]["y2"] = pt2.Y;
                                                    }

                                                }
                                            }



                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                if (dt1.Rows[i]["x1"] != DBNull.Value && dt1.Rows[i]["y1"] != DBNull.Value && dt1.Rows[i]["x2"] != DBNull.Value && dt1.Rows[i]["y2"] != DBNull.Value)
                                                {
                                                    double x1 = Convert.ToDouble(dt1.Rows[i]["x1"]);
                                                    double y1 = Convert.ToDouble(dt1.Rows[i]["y1"]);
                                                    double x2 = Convert.ToDouble(dt1.Rows[i]["x2"]);
                                                    double y2 = Convert.ToDouble(dt1.Rows[i]["y2"]);
                                                    Point3d pt1 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x1, y1, 0), Vector3d.ZAxis, false);
                                                    Point3d pt2 = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x2, y2, 0), Vector3d.ZAxis, false);



                                                    double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt1);
                                                    double param2 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt2);
                                                    if (param1 > _AGEN_mainform.Poly2D.EndParam)
                                                    {
                                                        param1 = _AGEN_mainform.Poly2D.EndParam;
                                                    }
                                                    if (param2 > _AGEN_mainform.Poly2D.EndParam)
                                                    {
                                                        param2 = _AGEN_mainform.Poly2D.EndParam;
                                                    }
                                                    Polyline poly_heavy_wall = Functions.get_part_of_poly(_AGEN_mainform.Poly2D, param1, param2);



                                                    string layer1 = "0";
                                                    if (dt1.Rows[i]["layer"] != DBNull.Value)
                                                    {
                                                        layer1 = Convert.ToString(dt1.Rows[i]["layer"]);
                                                    }


                                                    if (layertable1.Has(layer1) == false)
                                                    {
                                                        Functions.Creaza_layer(layer1, ci, true);
                                                        ++ci;
                                                        if (ci == 7) ci = 1;
                                                    }
                                                    poly_heavy_wall.Layer = layer1;
                                                    BTrecord.AppendEntity(poly_heavy_wall);
                                                    Trans1.AddNewlyCreatedDBObject(poly_heavy_wall, true);
                                                }

                                            }
                                        }
                                        #endregion

                                    }
                                }
                            }
                        }


                        if (dt1.Rows.Count > 0 && _AGEN_mainform.COUNTRY == "USA")
                        {
                            _AGEN_mainform.Poly3D.Erase();
                        }
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            set_enable_true();

            this.MdiParent.WindowState = FormWindowState.Normal;
        }

        private void Button_verify_order_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect the blocks:";
                        Prompt_rez.SingleOnly = false;
                        this.MdiParent.WindowState = FormWindowState.Minimized;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);

                        if (Rezultat1.Status == PromptStatus.OK)
                        {
                            System.Data.DataTable dt1 = new System.Data.DataTable();
                            dt1.Columns.Add("BlockName", typeof(string));
                            dt1.Columns.Add("Layer", typeof(string));
                            dt1.Columns.Add("Sta1", typeof(double));
                            dt1.Columns.Add("Sta2", typeof(double));
                            dt1.Columns.Add("Sta", typeof(double));
                            dt1.Columns.Add("Length", typeof(double));
                            dt1.Columns.Add("CalculatedLength", typeof(double));
                            dt1.Columns.Add("Reference_Sta", typeof(double));
                            dt1.Columns.Add("x", typeof(double));
                            dt1.Columns.Add("y", typeof(double));
                            dt1.Columns.Add("Distance1", typeof(double));
                            dt1.Columns.Add("ExcelFormula1", typeof(string));
                            dt1.Columns.Add("ExcelFormula2", typeof(string));

                            for (int i = 0; i < Rezultat1.Value.Count; ++i)
                            {
                                BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                                if (block1 != null)
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1]["BlockName"] = Functions.get_block_name(block1);
                                    dt1.Rows[dt1.Rows.Count - 1]["Layer"] = block1.Layer;
                                    dt1.Rows[dt1.Rows.Count - 1]["x"] = Math.Round(block1.Position.X, 3);
                                    dt1.Rows[dt1.Rows.Count - 1]["y"] = Math.Round(block1.Position.Y, 0);

                                    if (block1.AttributeCollection.Count > 0)
                                    {
                                        bool is_sta2 = false;
                                        for (int j = 0; j < block1.AttributeCollection.Count; ++j)
                                        {
                                            AttributeReference atr1 = Trans1.GetObject(block1.AttributeCollection[j], OpenMode.ForRead) as AttributeReference;
                                            if (atr1 != null)
                                            {
                                                string tag1 = atr1.Tag.ToLower();

                                                if (tag1 == "sta1")
                                                {
                                                    string val1 = Convert.ToString(atr1.TextString);
                                                    if (Functions.IsNumeric(val1.Replace("+", "")) == true)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1]["Sta1"] = Convert.ToDouble(val1.Replace("+", ""));
                                                    }
                                                }
                                                if (tag1 == "sta2")
                                                {
                                                    string val1 = Convert.ToString(atr1.TextString);
                                                    if (Functions.IsNumeric(val1.Replace("+", "")) == true)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1]["Sta2"] = Convert.ToDouble(val1.Replace("+", ""));
                                                        is_sta2 = true;
                                                    }
                                                }
                                                if (tag1 == "sta")
                                                {
                                                    string val1 = Convert.ToString(atr1.TextString);
                                                    if (Functions.IsNumeric(val1.Replace("+", "")) == true)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1]["Sta"] = Convert.ToDouble(val1.Replace("+", ""));
                                                    }
                                                }
                                                if (tag1 == "length" || tag1 == "len")
                                                {
                                                    string val1 = Convert.ToString(atr1.TextString);
                                                    if (Functions.IsNumeric(val1.Replace("'", "").Replace("m", "")) == true)
                                                    {
                                                        dt1.Rows[dt1.Rows.Count - 1]["Length"] = Convert.ToDouble(val1.Replace("'", "").Replace("m", ""));
                                                    }
                                                }
                                            }
                                        }

                                        double exist_dist1 = Functions.Get_Param_Value_block(block1, "Distance1");
                                        dt1.Rows[dt1.Rows.Count - 1]["Distance1"] = exist_dist1;

                                        if (is_sta2 == true)
                                        {
                                            int lr = 1;
                                            if (radioButton_rl.Checked == true) lr = -1;
                                            if (dt1.Rows[dt1.Rows.Count - 1]["Sta1"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[dt1.Rows.Count - 1]["Sta1"]).Replace("+", "")) == true &&
                                                dt1.Rows[dt1.Rows.Count - 1]["Sta2"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(dt1.Rows[dt1.Rows.Count - 1]["Sta2"]).Replace("+", "")) == true)
                                            {
                                                double sta1 = Convert.ToDouble(Convert.ToString(dt1.Rows[dt1.Rows.Count - 1]["Sta1"]).Replace("+", ""));
                                                double sta2 = Convert.ToDouble(Convert.ToString(dt1.Rows[dt1.Rows.Count - 1]["Sta2"]).Replace("+", ""));
                                                dt1.Rows[dt1.Rows.Count - 1]["CalculatedLength"] = Math.Abs(sta2 - sta1);
                                            }
                                            System.Data.DataTable dt2 = dt1.Copy();
                                            dt1.ImportRow(dt2.Rows[dt2.Rows.Count - 1]);
                                            dt2.Dispose();
                                            dt1.Rows[dt1.Rows.Count - 1]["x"] = Math.Round(block1.Position.X + lr * exist_dist1, 3);
                                            dt1.Rows[dt1.Rows.Count - 1]["Reference_Sta"] = dt1.Rows[dt1.Rows.Count - 1]["Sta2"];
                                            dt1.Rows[dt1.Rows.Count - 2]["Reference_Sta"] = dt1.Rows[dt1.Rows.Count - 2]["Sta1"];
                                            dt1.Rows[dt1.Rows.Count - 2]["Sta2"] = DBNull.Value;
                                            dt1.Rows[dt1.Rows.Count - 1]["Sta1"] = DBNull.Value;
                                        }
                                    }
                                }
                            }

                            for (int i = 0; i < dt1.Rows.Count; ++i)
                            {
                                if (dt1.Rows[i]["Reference_Sta"] == DBNull.Value)
                                {
                                    if (dt1.Rows[i]["Sta"] != DBNull.Value) dt1.Rows[i]["Reference_Sta"] = dt1.Rows[i]["Sta"];
                                    if (dt1.Rows[i]["Sta1"] != DBNull.Value) dt1.Rows[i]["Reference_Sta"] = dt1.Rows[i]["Sta1"];
                                    if (dt1.Rows[i]["Sta2"] != DBNull.Value) dt1.Rows[i]["Reference_Sta"] = dt1.Rows[i]["Sta2"];
                                }
                            }

                            for (int i = dt1.Rows.Count - 1; i >= 0; --i)
                            {
                                if (dt1.Rows[i]["Reference_Sta"] == DBNull.Value)
                                {
                                    dt1.Rows[i].Delete();
                                }
                            }

                            dt1 = Sort_data_table(dt1, "Reference_Sta");
                            string formula_excel = "IF(J3=J2,IF(I3>I2,1,0),";
                            if (radioButton_lr.Checked == true) formula_excel = "IF(J3=J2,IF(I3<I2,1,0),";
                            dt1.Rows[1]["ExcelFormula1"] = formula_excel + Convert.ToString((char)34) + Convert.ToString((char)34) + ")";
                            dt1.Rows[1]["ExcelFormula2"] = "IF(F3>0,IF(ROUND(F3,2)=ROUND(G3,2),0,1)," + Convert.ToString((char)34) + Convert.ToString((char)34) + ")";
                            Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.MdiParent.WindowState = FormWindowState.Normal;
            this.MdiParent.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - this.MdiParent.Width) / 2, (Screen.PrimaryScreen.WorkingArea.Height - this.MdiParent.Height) / 2);
            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();
        }

        System.Data.DataTable Sort_data_table(System.Data.DataTable dt1, string Column1)
        {
            System.Data.DataTable Data_table_temp = new System.Data.DataTable();
            if (dt1 != null)
            {
                if (dt1.Rows.Count > 0)
                {
                    if (dt1.Columns.Contains(Column1) == true)
                    {
                        dt1.DefaultView.Sort = Column1 + " ASC";
                        Data_table_temp = dt1.DefaultView.ToTable();
                    }
                }
            }
            return Data_table_temp;

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

        private void button_P2C_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            List<ObjectId> lista_objid = new List<ObjectId>();
            List<string> lista_chain = new List<string>();

            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;


                        string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                        {
                            ProjFolder = ProjFolder + "\\";
                        }
                        if (System.IO.Directory.Exists(ProjFolder) == true)
                        {
                            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                            if (System.IO.File.Exists(fisier_cl) == true)
                            {

                                if (_AGEN_mainform.dt_centerline == null)
                                {
                                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                                }

                                System.Data.DataTable dt_cl = _AGEN_mainform.dt_centerline;

                                Polyline poly1 = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                                this.MdiParent.WindowState = FormWindowState.Minimized;
                                bool repeat = true;

                                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                Functions.Create_mleader_object_data_table();

                                do
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify point:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {

                                        if (lista_chain.Count > 0)
                                        {
                                            Functions.Append_object_data_to_ODXXX(lista_objid, _AGEN_mainform.current_segment, lista_chain);
                                        }

                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Trans1.Commit();
                                        this.MdiParent.WindowState = FormWindowState.Normal;
                                        repeat = false;
                                    }

                                    Point3d pt1 = Point_res1.Value;
                                    Point3d pt2d = poly1.GetClosestPointTo(new Point3d(pt1.X, pt1.Y, 0), Vector3d.ZAxis, false);
                                    double dist2d = poly1.GetDistAtPoint(pt2d);
                                    double b1 = -1.23456;

                                    double calc_sta = Math.Round(Functions.get_stationCSF_from_point(poly1, pt2d, dist2d, dt_cl, ref b1), _AGEN_mainform.round1);
                                    double texth = 2;
                                    if (Functions.IsNumeric(textBox_text_height.Text) == true) texth = Convert.ToDouble(textBox_text_height.Text);

                                    MLeader mlead1 = Functions.creaza_mleader(new Point3d(pt2d.X, pt2d.Y, 0), Functions.Get_chainage_from_double(calc_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1), texth, texth, texth / 2, texth / 2, texth / 2, 0.1, _AGEN_mainform.layer_no_plot);

                                    lista_objid.Add(mlead1.ObjectId);
                                    lista_chain.Add(Functions.Get_chainage_from_double(calc_sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1));

                                    Trans1.TransactionManager.QueueForGraphicsFlush();

                                } while (repeat == true);
                            }
                        }

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

        private void button_read_band_to_xl_Click(object sender, EventArgs e)
        {
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            List<string> lista_mat = new List<string>();
            List<string> lista_xing = new List<string>();
            List<string> lista_sa = new List<string>();

            if (comboBox_mat1.Text != "") lista_mat.Add(comboBox_mat1.Text);
            if (comboBox_mat2.Text != "") lista_mat.Add(comboBox_mat2.Text);
            if (comboBox_mat3.Text != "") lista_mat.Add(comboBox_mat3.Text);
            if (comboBox_mat4.Text != "") lista_mat.Add(comboBox_mat4.Text);
            if (comboBox_mat5.Text != "") lista_mat.Add(comboBox_mat5.Text);
            if (comboBox_mat6.Text != "") lista_mat.Add(comboBox_mat6.Text);
            if (comboBox_mat7.Text != "") lista_mat.Add(comboBox_mat7.Text);
            if (comboBox_mat8.Text != "") lista_mat.Add(comboBox_mat8.Text);

            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = Trans1.GetObject(ThisDrawing.Database.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead) as BlockTable;

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat1;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_rez = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        Prompt_rez.MessageForAdding = "\nSelect all objects:";
                        Prompt_rez.SingleOnly = false;
                        this.MdiParent.WindowState = FormWindowState.Minimized;
                        Rezultat1 = ThisDrawing.Editor.GetSelection(Prompt_rez);
                        if (Rezultat1.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            this.MdiParent.WindowState = FormWindowState.Normal;
                            return;
                        }
                        this.MdiParent.WindowState = FormWindowState.Normal;

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        dt1.Columns.Add(col_handle, typeof(string));
                        dt1.Columns.Add(col_blockname, typeof(string));
                        dt1.Columns.Add(col_sta, typeof(double));
                        dt1.Columns.Add(col_x, typeof(double));
                        dt1.Columns.Add(col_y, typeof(double));
                        dt1.Columns.Add(col_visibility, typeof(string));
                        dt1.Columns.Add(col_dist1, typeof(double));
                        dt1.Columns.Add(col_dist2, typeof(double));
                        dt1.Columns.Add(col_dist3, typeof(double));
                        dt1.Columns.Add(col_dist4, typeof(double));

                        for (int i = 0; i < Rezultat1.Value.Count; ++i)
                        {
                            BlockReference block1 = Trans1.GetObject(Rezultat1.Value[i].ObjectId, OpenMode.ForRead) as BlockReference;
                            if (block1 != null)
                            {
                                string blockname = Functions.get_block_name(block1);
                                if (lista_mat.Contains(blockname) == true || lista_xing.Contains(blockname) == true || lista_sa.Contains(blockname) == true)
                                {
                                    dt1.Rows.Add();
                                    dt1.Rows[dt1.Rows.Count - 1][col_handle] = block1.ObjectId.Handle.Value.ToString();
                                    dt1.Rows[dt1.Rows.Count - 1][col_blockname] = blockname;
                                    dt1.Rows[dt1.Rows.Count - 1][col_x] = block1.Position.X;
                                    dt1.Rows[dt1.Rows.Count - 1][col_y] = block1.Position.Y;
                                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attColl = block1.AttributeCollection;
                                    if (attColl.Count > 0)
                                    {
                                        foreach (ObjectId id1 in attColl)
                                        {
                                            DBObject ent = Trans1.GetObject(id1, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                                            if (ent is AttributeReference)
                                            {
                                                AttributeReference atr1 = ent as AttributeReference;
                                                string column1 = atr1.Tag;
                                                if (dt1.Columns.Contains(atr1.Tag) == false)
                                                {
                                                    dt1.Columns.Add(atr1.Tag, typeof(string));
                                                }
                                                else
                                                {
                                                    //string new_col = atr1.Tag;
                                                    //int dupl = 1;
                                                    //string col1 = "";
                                                    //do
                                                    //{
                                                    //    col1 = new_col + "_duplicate" + dupl.ToString();
                                                    //    ++dupl;
                                                    //} while (dt1.Columns.Contains(col1) == true);


                                                    //dt1.Columns.Add(col1, typeof(string));
                                                    //column1 = col1;

                                                }
                                                if (atr1.IsMTextAttribute == false)
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][column1] = atr1.TextString;
                                                }
                                                else
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][column1] = atr1.MTextAttribute.Contents.Replace("\\P", "\r\n");
                                                }

                                                if (atr1.Tag == "STA")
                                                {
                                                    string val1 = atr1.TextString;
                                                    if (Functions.IsNumeric(val1.Replace("+", "")) == true)
                                                    {
                                                        double nr1 = Convert.ToDouble(val1.Replace("+", ""));
                                                        dt1.Rows[dt1.Rows.Count - 1][col_sta] = nr1;
                                                    }
                                                }
                                                if (atr1.Tag == "STA1" && dt1.Rows[dt1.Rows.Count - 1][col_sta] == DBNull.Value)
                                                {
                                                    string val1 = atr1.TextString;
                                                    if (Functions.IsNumeric(val1.Replace("+", "")) == true)
                                                    {
                                                        double nr1 = Convert.ToDouble(val1.Replace("+", ""));
                                                        dt1.Rows[dt1.Rows.Count - 1][col_sta] = nr1;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (block1.IsDynamicBlock == true)
                                    {
                                        using (DynamicBlockReferencePropertyCollection pc = block1.DynamicBlockReferencePropertyCollection)
                                        {
                                            foreach (DynamicBlockReferenceProperty prop in pc)
                                            {
                                                if (prop.PropertyName == "Visibility1")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][col_visibility] = Convert.ToString(prop.Value);
                                                }
                                                else if (prop.PropertyName == "Distance1")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][col_dist1] = Convert.ToDouble(prop.Value);
                                                }
                                                else if (prop.PropertyName == "Distance2")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][col_dist2] = Convert.ToDouble(prop.Value);
                                                }
                                                else if (prop.PropertyName == "Distance3")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][col_dist3] = Convert.ToDouble(prop.Value);
                                                }
                                                else if (prop.PropertyName == "Distance4")
                                                {
                                                    dt1.Rows[dt1.Rows.Count - 1][col_dist4] = Convert.ToDouble(prop.Value);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (dt1.Columns.Contains("STA1") == true) dt1.Columns["STA1"].SetOrdinal(3);
                        if (dt1.Columns.Contains("STA") == true)
                        {
                            if (dt1.Columns.Contains("STA1") == true)
                            {
                                dt1.Columns["STA"].SetOrdinal(4);
                            }
                            else
                            {
                                dt1.Columns["STA"].SetOrdinal(3);
                            }
                        }
                        if (dt1.Columns.Count > 5 && dt1.Columns.Contains("STA2") == true)
                        {
                            dt1.Columns["STA2"].SetOrdinal(5);
                        }
                        if (dt1.Columns.Count > 6 && dt1.Columns.Contains("LEN") == true)
                        {
                            dt1.Columns["LEN"].SetOrdinal(6);
                        }
                        if (dt1.Columns.Count > 7 && dt1.Columns.Contains("MAT") == true)
                        {
                            dt1.Columns["MAT"].SetOrdinal(7);
                        }
                        dt1.Columns[col_x].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_y].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_visibility].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_dist1].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_dist2].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_dist3].SetOrdinal(dt1.Columns.Count - 1);
                        dt1.Columns[col_dist4].SetOrdinal(dt1.Columns.Count - 1);
                        if (dt1.Columns.Count > 15)
                        {
                            List<string> lista1 = new List<string>();
                            for (int j = 8; j < dt1.Columns.Count - 7; ++j)
                            {
                                lista1.Add(dt1.Columns[j].ColumnName);
                            }
                            lista1.Sort();
                            for (int i = 0; i < lista1.Count; ++i)
                            {
                                dt1.Columns[lista1[i]].SetOrdinal(dt1.Columns.Count - 8);
                            }
                        }
                        if (radioButton_rl.Checked == true)
                        {
                            dt1 = Functions.Sort_data_table_DESC(dt1, col_x);
                        }
                        else
                        {
                            dt1 = Functions.Sort_data_table(dt1, col_x);
                        }
                        Functions.Transfer_datatable_to_new_excel_spreadsheet_formated_general(dt1, "EngBand");
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

        private void button_calc_length_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;

                try
                {
                    Microsoft.Office.Interop.Excel.Application Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook2.Worksheets)
                        {
                            if (W2.Name == "EngBand")
                            {
                                W1 = W2;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    return;
                }
                if (W1 != null)
                {
                    try
                    {
                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        string last_letter = Functions.get_excel_column_letter(100);
                        Microsoft.Office.Interop.Excel.Range Column_row = W1.Range["A1:" + last_letter + "1"];
                        object[,] Values_for_columns = new object[1, 100];
                        Values_for_columns = Column_row.Value2;

                        string col_sta1 = "0";
                        string col_sta2 = "0";
                        string col_len = "0";

                        for (int i = 1; i <= Values_for_columns.Length; ++i)
                        {
                            object val = Values_for_columns[1, i];
                            if (val != null)
                            {
                                string col_name = Convert.ToString(val);
                                if (col_name.ToLower() == "sta1")
                                {
                                    if (dt1.Columns.Contains("sta1") == false)
                                    {
                                        dt1.Columns.Add("sta1", typeof(double));
                                        col_sta1 = Functions.get_excel_column_letter(i);
                                    }
                                }

                                if (col_name.ToLower() == "sta2")
                                {
                                    if (dt1.Columns.Contains("sta2") == false)
                                    {
                                        dt1.Columns.Add("sta2", typeof(double));
                                        col_sta2 = Functions.get_excel_column_letter(i);
                                    }
                                }

                                if (col_name.ToLower() == "length" || col_name.ToLower() == "len")
                                {
                                    if (dt1.Columns.Contains("length") == false)
                                    {
                                        dt1.Columns.Add("length", typeof(double));
                                        col_len = Functions.get_excel_column_letter(i);
                                    }
                                }
                                if (col_sta1 != "0" && col_sta2 != "0" && col_len != "0")
                                {
                                    i = Values_for_columns.Length + 1;
                                }

                            }
                            else
                            {
                                i = Values_for_columns.Length + 1;
                            }
                        }


                        if (col_sta1 != "0" && col_sta2 != "0" && col_len != "0")
                        {
                            Microsoft.Office.Interop.Excel.Range range2 = W1.Range[col_sta1 + 2.ToString() + ":" + col_sta1 + "30001"];
                            Microsoft.Office.Interop.Excel.Range range3 = W1.Range[col_sta2 + 2.ToString() + ":" + col_sta2 + "30001"];
                            object[,] values_for_check2 = new object[30000, 1];
                            values_for_check2 = range2.Value2;
                            object[,] values_for_check3 = new object[30000, 1];
                            values_for_check3 = range3.Value2;

                            bool is_data = false;
                            for (int i = 1; i <= values_for_check2.Length; ++i)
                            {
                                object Valoare2 = values_for_check2[i, 1];
                                object Valoare3 = values_for_check3[i, 1];
                                if (Valoare2 != null && Valoare3 != null)
                                {
                                    dt1.Rows.Add();
                                    is_data = true;
                                }
                                else
                                {
                                    i = values_for_check2.Length + 1;
                                }
                            }

                            if (is_data == true)
                            {
                                int NrR = dt1.Rows.Count;
                                int NrC = dt1.Columns.Count;
                                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[col_sta1 + "2:" + col_sta1 + Convert.ToString(NrR + 1)];
                                Microsoft.Office.Interop.Excel.Range range11 = W1.Range[col_sta2 + "2:" + col_sta2 + Convert.ToString(NrR + 1)];
                                Microsoft.Office.Interop.Excel.Range range111 = W1.Range[col_len + "2:" + col_len + Convert.ToString(NrR + 1)];
                                object[,] values_with_data1 = new object[NrR - 1, 1];
                                object[,] values_with_data2 = new object[NrR - 1, 1];

                                values_with_data1 = range1.Value2;
                                values_with_data2 = range11.Value2;

                                for (int i = 0; i < dt1.Rows.Count; ++i)
                                {
                                    object Valoare1 = values_with_data1[i + 1, 1];
                                    object Valoare2 = values_with_data2[i + 1, 1];
                                    if (Valoare1 != null)
                                    {
                                        string valu1 = Convert.ToString(Valoare1).Replace("+", "");
                                        if (Functions.IsNumeric(valu1) == true)
                                        {
                                            dt1.Rows[i][0] = Convert.ToDouble(valu1);
                                        }
                                        else
                                        {
                                            dt1.Rows[i][0] = DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        dt1.Rows[i][0] = DBNull.Value;
                                    }

                                    if (Valoare2 != null)
                                    {
                                        string valu2 = Convert.ToString(Valoare2).Replace("+", "");
                                        if (Functions.IsNumeric(valu2) == true)
                                        {
                                            dt1.Rows[i][1] = Convert.ToDouble(valu2);
                                        }
                                        else
                                        {
                                            dt1.Rows[i][1] = DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        dt1.Rows[i][1] = DBNull.Value;
                                    }
                                }

                                for (int i = 0; i < dt1.Rows.Count; ++i)
                                {
                                    if (dt1.Rows[i][0] != DBNull.Value && dt1.Rows[i][1] != DBNull.Value)
                                    {
                                        double sta1 = Convert.ToDouble(dt1.Rows[i][0]);
                                        double sta2 = Convert.ToDouble(dt1.Rows[i][1]);

                                        dt1.Rows[i][2] = sta2 - sta1;
                                    }
                                }

                                object[,] values_sta1 = new object[NrR, 1];
                                object[,] values_sta2 = new object[NrR, 1];
                                object[,] values_len = new object[NrR, 1];

                                for (int i = 0; i < dt1.Rows.Count; ++i)
                                {
                                    if (dt1.Rows[i][0] != DBNull.Value && dt1.Rows[i][1] != DBNull.Value)
                                    {
                                        double sta1 = Convert.ToDouble(dt1.Rows[i][0]);
                                        double sta2 = Convert.ToDouble(dt1.Rows[i][1]);
                                        double len = Convert.ToDouble(dt1.Rows[i][2]);

                                        values_sta1[i, 0] = Functions.Get_chainage_from_double(sta1, "m", 1);
                                        values_sta2[i, 0] = Functions.Get_chainage_from_double(sta2, "m", 1);
                                        values_len[i, 0] = Functions.Get_String_Rounded(len, 1);
                                    }
                                }
                                range1.Value2 = values_sta1;
                                range11.Value2 = values_sta2;
                                range111.NumberFormat = "@";
                                range111.Value2 = values_len;
                                set_enable_true();
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
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
        }

        private void button_update_band_Click(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Worksheet W1 = null;
                try
                {
                    Microsoft.Office.Interop.Excel.Application Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        foreach (Microsoft.Office.Interop.Excel.Worksheet W2 in Workbook2.Worksheets)

                        {
                            if (W2.Name == "EngBand")
                            {
                                W1 = W2;
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    return;
                }
                if (W1 != null)
                {
                    try
                    {

                        int nr_col = 0;

                        System.Data.DataTable dt1 = new System.Data.DataTable();
                        string last_letter = Functions.get_excel_column_letter(100);
                        Microsoft.Office.Interop.Excel.Range Column_row = W1.Range["A1:" + last_letter + "1"];
                        object[,] Values_for_columns = new object[1, 100];
                        Values_for_columns = Column_row.Value2;

                        for (int i = 1; i <= Values_for_columns.Length; ++i)
                        {
                            object val = Values_for_columns[1, i];
                            if (val != null)
                            {
                                string col_name = Convert.ToString(val);
                                if (dt1.Columns.Contains(col_name) == false)
                                {
                                    if (col_name == col_x || col_name == col_y || col_name == col_dist1 || col_name == col_dist2 || col_name == col_dist3 || col_name == col_dist4)
                                    {
                                        dt1.Columns.Add(col_name, typeof(double));
                                    }
                                    else
                                    {
                                        dt1.Columns.Add(col_name, typeof(string));
                                    }
                                    ++nr_col;
                                }
                            }
                            else
                            {
                                i = Values_for_columns.Length + 1;
                            }
                        }

                        if (nr_col > 0)
                        {
                            string Col1 = "A";
                            Microsoft.Office.Interop.Excel.Range range2 = W1.Range[Col1 + 2.ToString() + ":" + Col1 + "30001"];
                            object[,] values_for_check = new object[30000, 1];
                            values_for_check = range2.Value2;

                            bool is_data = false;
                            for (int i = 1; i <= values_for_check.Length; ++i)
                            {
                                object Valoare2 = values_for_check[i, 1];
                                if (Valoare2 != null)
                                {
                                    dt1.Rows.Add();
                                    is_data = true;
                                }
                                else
                                {
                                    i = values_for_check.Length + 1;
                                }
                            }

                            if (is_data == true)
                            {
                                int NrR = dt1.Rows.Count;
                                int NrC = dt1.Columns.Count;

                                Microsoft.Office.Interop.Excel.Range range1 = W1.Range[W1.Cells[2, 1], W1.Cells[NrR + 2 - 1, NrC]];
                                object[,] values_with_data = new object[NrR - 1, NrC - 1];
                                values_with_data = range1.Value2;
                                for (int i = 0; i < dt1.Rows.Count; ++i)
                                {
                                    for (int j = 0; j < dt1.Columns.Count; ++j)
                                    {
                                        object Valoare = values_with_data[i + 1, j + 1];
                                        if (Valoare != null)
                                        {
                                            dt1.Rows[i][j] = Convert.ToString(Valoare).Replace("\r\n", "\\P").Replace("\n", "\\P");
                                        }
                                        else
                                        {
                                            dt1.Rows[i][j] = DBNull.Value;
                                        }
                                    }
                                }

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

                                            BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                            for (int i = 0; i < dt1.Rows.Count; ++i)
                                            {
                                                ObjectId id1 = ObjectId.Null;
                                                if (dt1.Rows[i][col_handle] != DBNull.Value)
                                                {
                                                    string handle1 = Convert.ToString(dt1.Rows[i][col_handle]);

                                                    id1 = Functions.GetObjectId(ThisDrawing.Database, handle1);

                                                    BlockReference block1 = Trans1.GetObject(id1, OpenMode.ForWrite) as BlockReference;
                                                    bool delete_block = false;
                                                    if (dt1.Rows[i][col_blockname] != DBNull.Value)
                                                    {
                                                        string bn = Convert.ToString(dt1.Rows[i][col_blockname]);
                                                        if (bn.ToUpper() == "DELETE") delete_block = true;
                                                    }

                                                    double x = block1.Position.X;
                                                    double y = block1.Position.Y;
                                                    double z = block1.Position.Z;

                                                    for (int j = 0; j < dt1.Columns.Count; ++j)
                                                    {
                                                        string column_name = dt1.Columns[j].ColumnName;
                                                        if (column_name == "X")
                                                        {
                                                            x = Convert.ToDouble(dt1.Rows[i][j]);
                                                        }
                                                        if (column_name == "Y")
                                                        {
                                                            y = Convert.ToDouble(dt1.Rows[i][j]);
                                                        }
                                                    }
                                                    if (delete_block == false)
                                                    {
                                                        block1.Position = new Point3d(x, y, z);
                                                    }

                                                    if (delete_block == false && block1.AttributeCollection.Count > 0)
                                                    {
                                                        foreach (ObjectId id2 in block1.AttributeCollection)
                                                        {
                                                            DBObject ent = Trans1.GetObject(id2, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
                                                            if (ent is AttributeReference)
                                                            {
                                                                AttributeReference atr1 = ent as AttributeReference;

                                                                for (int j = 0; j < dt1.Columns.Count; ++j)
                                                                {
                                                                    string column_name = dt1.Columns[j].ColumnName;

                                                                    if (atr1.Tag == column_name)
                                                                    {
                                                                        string val1 = "";
                                                                        if (dt1.Rows[i][j] != DBNull.Value)
                                                                        {
                                                                            val1 = Convert.ToString(dt1.Rows[i][j]);
                                                                        }

                                                                        if (atr1.IsMTextAttribute == true)
                                                                        {
                                                                            atr1.TextString = val1;
                                                                            atr1.UpdateMTextAttribute();
                                                                        }
                                                                        else
                                                                        {
                                                                            atr1.TextString = val1;
                                                                        }

                                                                        j = dt1.Columns.Count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    if (delete_block == true)
                                                    {
                                                        block1.Erase();
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

                                Editor1.SetImpliedSelection(Empty_array);
                                Editor1.WriteMessage("\nCommand:");
                                set_enable_true();
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
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
        }

        private void comboBox_mat1_DropDown(object sender, EventArgs e)
        {
            if (checkBox_load_dollar.Checked == false)
            {
                Functions.Incarca_existing_Blocks_to_combobox(sender as ComboBox);
            }
            else
            {
                Functions.Incarca_existing_Blocks_with_dollar_to_combobox(sender as ComboBox);
            }
        }

        private void comboBox_ws1_DropDown(object sender, EventArgs e)
        {
            Functions.Load_opened_worksheets_to_combobox(comboBox_ws1);
            for (int i = 0; i < comboBox_ws1.Items.Count; ++i)
            {
                string item1 = Convert.ToString(comboBox_ws1.Items[i]);
                if (item1.ToLower().Contains("[materials]") == true)
                {
                    comboBox_ws1.SelectedIndex = i;
                    i = comboBox_ws1.Items.Count;
                }
            }
        }

        private void button_arc_to_lines_Click(object sender, EventArgs e)
        {
            if (Functions.IsNumeric(textBox_min_dist.Text) == false)
            {
                MessageBox.Show("min dist not a number");
                return;
            }
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
                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        string ln = "__PL";
                        Functions.Creaza_layer(ln, 5, true);


                        Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_arc;
                        Autodesk.AutoCAD.EditorInput.PromptSelectionOptions prompt_arc = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                        prompt_arc.MessageForAdding = "\nSelect arcs:";
                        prompt_arc.SingleOnly = false;
                        this.MdiParent.WindowState = FormWindowState.Minimized;

                        Rezultat_arc = ThisDrawing.Editor.GetSelection(prompt_arc);
                        this.MdiParent.WindowState = FormWindowState.Normal;
                        if (Rezultat_arc.Status != PromptStatus.OK)
                        {

                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }


                        for (int k = 0; k < Rezultat_arc.Value.Count; ++k)
                        {
                            Arc arc1 = Trans1.GetObject(Rezultat_arc.Value[k].ObjectId, OpenMode.ForRead) as Arc;

                            if (arc1 != null)
                            {
                                Polyline poly1 = new Polyline();
                                poly1.Layer = ln;

                                Point3d pt_cen = arc1.Center;
                                double r1 = arc1.Radius;
                                Point3d p1 = arc1.StartPoint;
                                Point3d p2 = arc1.EndPoint;
                                double d1 = Convert.ToDouble(textBox_min_dist.Text);

                                double l = arc1.Length;

                                int no_segm = Convert.ToInt32(Math.Floor(l / d1));

                                poly1.AddVertexAt(0, new Point2d(p1.X, p1.Y), 0, 0, 0);

                                for (int i = 0; i < no_segm; ++i)
                                {
                                    Polyline poly_start = new Polyline();
                                    poly_start.AddVertexAt(0, new Point2d(pt_cen.X, pt_cen.Y), 0, 0, 0);
                                    poly_start.AddVertexAt(1, new Point2d(p1.X, p1.Y), 0, 0, 0);

                                    double rot1 = 2 * Math.Asin(0.5 * d1 / r1);

                                    poly_start.TransformBy(Matrix3d.Rotation(rot1, Vector3d.ZAxis, pt_cen));
                                    poly1.AddVertexAt(i + 1, new Point2d(poly_start.EndPoint.X, poly_start.EndPoint.Y), 0, 0, 0);
                                    p1 = new Point3d(poly_start.EndPoint.X, poly_start.EndPoint.Y, 0);

                                }

                                poly1.AddVertexAt(poly1.NumberOfVertices, new Point2d(p2.X, p2.Y), 0, 0, 0);




                                BTrecord.AppendEntity(poly1);
                                Trans1.AddNewlyCreatedDBObject(poly1, true);


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

            Editor1.SetImpliedSelection(Empty_array);
            Editor1.WriteMessage("\nCommand:");
            set_enable_true();

        }

        private void button_project_cl_bends_to_profile_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        ObjectId[] Empty_array = null;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

                        set_enable_false();
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                                System.Data.DataTable dtH = _AGEN_mainform.dt_centerline;


                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_z = "Z";
                                string Col_Sta = "STA";
                                string Col_Deflection = "Deflection";
                                string Col_Defldms = "DeflectionDMS";
                                string Col_Side = "Side";
                                string Col_empty = "**";


                                make_first_line_invisible();


                                System.Data.DataTable dt2 = new System.Data.DataTable();

                                dt2.Columns.Add(Col_Sta, typeof(double));
                                dt2.Columns.Add(Col_z, typeof(double));
                                dt2.Columns.Add(Col_empty, typeof(string));
                                dt2.Columns.Add(Col_Deflection, typeof(double));
                                dt2.Columns.Add(Col_Defldms, typeof(string));
                                dt2.Columns.Add(Col_Side, typeof(string));
                                dt2.Columns.Add(Col_x, typeof(double));
                                dt2.Columns.Add(Col_y, typeof(double));
                                dt2.Columns.Add("Segment", typeof(string));


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
                                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                            if (W1 != null)
                                            {

                                                System.Data.DataTable dtV = new System.Data.DataTable();
                                                dtV.Columns.Add(Col_Sta, typeof(double));
                                                dtV.Columns.Add(Col_z, typeof(double));

                                                List<string> lista_col = new List<string>();
                                                List<string> lista_colxl = new List<string>();

                                                lista_col.Add(Col_Sta);
                                                lista_col.Add(Col_z);

                                                lista_colxl.Add(textBox_chainage.Text);
                                                lista_colxl.Add(textBox_col_z.Text);


                                                dtV = Functions.build_dt_from_excel(dtV, W1, start1, end1, lista_col, lista_colxl);



                                                if (dtV.Rows.Count > 0)
                                                {

                                                    for (int i = 0; i < dtH.Rows.Count; ++i)
                                                    {
                                                        if (dtH.Rows[i][_AGEN_mainform.Col_3DSta] != DBNull.Value)
                                                        {
                                                            double sta3d = Convert.ToDouble(dtH.Rows[i][_AGEN_mainform.Col_3DSta]);

                                                            if (i > 0 && i < dtH.Rows.Count - 1)
                                                            {
                                                                if (dtH.Rows[i - 1][_AGEN_mainform.Col_x] != DBNull.Value && dtH.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value && dtH.Rows[i + 1][_AGEN_mainform.Col_x] != DBNull.Value)
                                                                {
                                                                    if (dtH.Rows[i - 1][_AGEN_mainform.Col_y] != DBNull.Value && dtH.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value && dtH.Rows[i + 1][_AGEN_mainform.Col_y] != DBNull.Value)
                                                                    {
                                                                        double x1 = Convert.ToDouble(dtH.Rows[i - 1][_AGEN_mainform.Col_x]);
                                                                        double y1 = Convert.ToDouble(dtH.Rows[i - 1][_AGEN_mainform.Col_y]);
                                                                        double x2 = Convert.ToDouble(dtH.Rows[i][_AGEN_mainform.Col_x]);
                                                                        double y2 = Convert.ToDouble(dtH.Rows[i][_AGEN_mainform.Col_y]);
                                                                        double x3 = Convert.ToDouble(dtH.Rows[i + 1][_AGEN_mainform.Col_x]);
                                                                        double y3 = Convert.ToDouble(dtH.Rows[i + 1][_AGEN_mainform.Col_y]);

                                                                        double defl = Functions.Get_deflection_angle_as_double(x1, y1, x2, y2, x3, y3);
                                                                        string side = Functions.Get_deflection_side(x1, y1, x2, y2, x3, y3);
                                                                        string defldms = Functions.Get_deflection_angle_dms(x1, y1, x2, y2, x3, y3);

                                                                        dt2.Rows.Add();
                                                                        dt2.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_x] = x2;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_y] = y2;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_Sta] = Math.Round(sta3d, 3);
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_Deflection] = defl;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_Defldms] = defldms;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_Side] = side;
                                                                        dt2.Rows[dt2.Rows.Count - 1][Col_z] = calc_z(dtV, sta3d);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (dtH.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value && dtH.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                                                {
                                                                    double x2 = Convert.ToDouble(dtH.Rows[i][_AGEN_mainform.Col_x]);
                                                                    double y2 = Convert.ToDouble(dtH.Rows[i][_AGEN_mainform.Col_y]);
                                                                    dt2.Rows.Add();
                                                                    dt2.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                                    dt2.Rows[dt2.Rows.Count - 1][Col_x] = x2;
                                                                    dt2.Rows[dt2.Rows.Count - 1][Col_y] = y2;
                                                                    dt2.Rows[dt2.Rows.Count - 1][Col_Sta] = Math.Round(sta3d, 3);
                                                                    dt2.Rows[dt2.Rows.Count - 1][Col_z] = calc_z(dtV, sta3d);
                                                                }
                                                            }
                                                        }
                                                    }
                                                    DateTime datetime1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                                    string nume1 = System.Environment.UserName + "-" + datetime1.Month + "_" + datetime1.Day + "at" + datetime1.Hour + "h" + datetime1.Minute + "m";
                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt2, nume1);
                                                }
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
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private double calc_z(System.Data.DataTable dt1, double sta3d)
        {

            string Col_z = "Z";
            string Col_Sta = "STA";

            for (int j = 0; j < dt1.Rows.Count - 1; ++j)
            {
                double sta1 = -1.234;
                if (dt1.Rows[j][Col_Sta] != DBNull.Value)
                {
                    sta1 = Convert.ToDouble(dt1.Rows[j][Col_Sta]);
                }

                double z1 = -1.234;
                if (dt1.Rows[j][Col_z] != DBNull.Value)
                {
                    z1 = Convert.ToDouble(dt1.Rows[j][Col_z]);
                }

                double sta2 = -1.234;
                if (dt1.Rows[j + 1][Col_Sta] != DBNull.Value)
                {
                    sta2 = Convert.ToDouble(dt1.Rows[j + 1][Col_Sta]);
                }

                double z2 = -1.234;
                if (dt1.Rows[j + 1][Col_z] != DBNull.Value)
                {
                    z2 = Convert.ToDouble(dt1.Rows[j + 1][Col_z]);
                }

                if (sta1 != -1.234 && z1 != -1.234 && sta2 != -1.234 && z2 != -1.234)
                {
                    if (Math.Round(sta3d, 2) >= Math.Round(sta1, 2) && Math.Round(sta3d, 2) <= Math.Round(sta2, 2))
                    {
                        return Math.Round(z2 + ((sta2 - sta3d) * (z1 - z2) / (sta2 - sta1)), 3);
                    }
                }



            }

            return -10000000;
        }

        private void button_calc_2D_distance_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        ObjectId[] Empty_array = null;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        set_enable_false();

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                                Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                Polyline3d poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_z = "Z";
                                string Col_Sta2D = "2D distance";
                                string Col_offset = "2D Offset";
                                string Col_Sta3D = "3D Chainage";

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
                                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                            if (W1 != null)
                                            {
                                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                                dt1.Columns.Add(Col_x, typeof(double));
                                                dt1.Columns.Add(Col_y, typeof(double));
                                                dt1.Columns.Add(Col_z, typeof(double));
                                                dt1.Columns.Add(Col_Sta2D, typeof(double));
                                                dt1.Columns.Add(Col_Sta3D, typeof(double));
                                                dt1.Columns.Add(Col_offset, typeof(double));
                                                dt1.Columns.Add("2Dpoint", typeof(string));

                                                List<string> lista_col = new List<string>();
                                                List<string> lista_colxl = new List<string>();
                                                lista_col.Add(Col_x);
                                                lista_col.Add(Col_y);
                                                lista_colxl.Add(textBox_col_e.Text);
                                                lista_colxl.Add(textBox_col_n.Text);

                                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, lista_col, lista_colxl);
                                                dt_errors = new System.Data.DataTable();
                                                dt_errors.Columns.Add("Survey File Value", typeof(double));
                                                dt_errors.Columns.Add("Calculated Value", typeof(string));
                                                dt_errors.Columns.Add("Error Type", typeof(string));

                                                if (dt1.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        double x = -1.234;
                                                        if (dt1.Rows[i][Col_x] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_x])) == true)
                                                        {
                                                            x = Convert.ToDouble(dt1.Rows[i][Col_x]);
                                                        }
                                                        double y = -1.234;
                                                        if (dt1.Rows[i][Col_y] != DBNull.Value &&
                                                            Functions.IsNumeric(Convert.ToString(dt1.Rows[i][Col_y])) == true)
                                                        {
                                                            y = Convert.ToDouble(dt1.Rows[i][Col_y]);
                                                        }
                                                        Point3d pt2d = poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                                        double dist2d = poly2D.GetDistAtPoint(pt2d);
                                                        if (x != -1.234 && y != -1.234)
                                                        {
                                                            double b1 = -1.23456;
                                                            double sta = Functions.get_stationCSF_from_point(poly2D, pt2d, dist2d, _AGEN_mainform.dt_centerline, ref b1);

                                                            dt1.Rows[i][Col_Sta2D] = dist2d;
                                                            dt1.Rows[i][Col_Sta3D] = sta;

                                                            dt1.Rows[i]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                            Point3d pt_on_poly = poly2D.GetClosestPointTo(new Point3d(x, y, poly2D.Elevation), Vector3d.ZAxis, false);
                                                            double param1 = poly2D.GetParameterAtPoint(pt_on_poly);
                                                            if (param1 > poly3D.EndParam) param1 = poly3D.EndParam;
                                                            Point3d point_for_z = poly3D.GetPointAtParameter(param1);
                                                            dt1.Rows[i][Col_z] = point_for_z.Z;
                                                            double offset1 = Math.Pow(Math.Pow(x - pt_on_poly.X, 2) + Math.Pow(y - pt_on_poly.Y, 2), 0.5);
                                                            dt1.Rows[i][Col_offset] = offset1;
                                                            dt1.Rows[i][Col_x] = x;
                                                            dt1.Rows[i][Col_y] = y;
                                                            if (b1 != -1.23456)
                                                            {
                                                                if (dt1.Columns.Contains("BacK Station") == false) dt1.Columns.Add("BacK Station", typeof(double));
                                                                if (dt1.Columns.Contains("Ahead Station") == false) dt1.Columns.Add("Ahead Station", typeof(double));
                                                                dt1.Rows[i]["Ahead Station"] = sta;
                                                                dt1.Rows[i]["BacK Station"] = b1;
                                                            }
                                                        }
                                                        else
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric x or y";
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = "See Row " + (start1 + i).ToString();
                                                        }
                                                    }
                                                    transfer_errors_to_panel(dt_errors);
                                                    dt1.Columns.Add("Segment", typeof(string));
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        dt1.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                    }
                                                    string nume1 = System.DateTime.Now.Hour + "-" + System.DateTime.Now.Minute;
                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1);
                                                }
                                            }
                                        }
                                    }
                                }
                                poly3D.Erase();
                                Trans1.Commit();
                            }
                        }
                        Editor1.SetImpliedSelection(Empty_array);
                        Editor1.WriteMessage("\nCommand:");
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }


        private void button_3D_2D_Click(object sender, EventArgs e)
        {
            int start1 = 0;
            int end1 = 0;

            if (Functions.IsNumeric(textBox_row_start.Text) == true)
            {
                start1 = Convert.ToInt32(textBox_row_start.Text);
            }

            if (Functions.IsNumeric(textBox_row_end.Text) == true)
            {
                end1 = Convert.ToInt32(textBox_row_end.Text);
            }

            if (start1 <= 0 || end1 <= 0 || start1 > end1)
            {
                MessageBox.Show("specify the start/end row!");
                return;
            }

            try
            {
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                    if (System.IO.File.Exists(fisier_cl) == true)
                    {
                        ObjectId[] Empty_array = null;
                        Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                        Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                        Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                        set_enable_false();

                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

                                Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                                Polyline3d poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                string Col_x = "X";
                                string Col_y = "Y";
                                string Col_z = "Z";
                                string Col_Sta2D = "2D distance";
                                string Col_Sta3D = "3D Chainage";

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
                                            Microsoft.Office.Interop.Excel.Worksheet W1 = Functions.Get_opened_worksheet_from_Excel_by_name(filename, sheet_name);
                                            if (W1 != null)
                                            {
                                                System.Data.DataTable dt1 = new System.Data.DataTable();
                                                dt1.Columns.Add(Col_x, typeof(double));
                                                dt1.Columns.Add(Col_y, typeof(double));
                                                dt1.Columns.Add(Col_z, typeof(double));
                                                dt1.Columns.Add(Col_Sta2D, typeof(double));
                                                dt1.Columns.Add(Col_Sta3D, typeof(double));
                                                dt1.Columns.Add("2Dpoint", typeof(string));

                                                List<string> lista_col = new List<string>();
                                                List<string> lista_colxl = new List<string>();
                                                lista_col.Add(Col_Sta3D);

                                                lista_colxl.Add(textBox_chainage.Text);

                                                dt1 = Functions.build_dt_from_excel(dt1, W1, start1, end1, lista_col, lista_colxl);
                                                dt_errors = new System.Data.DataTable();
                                                dt_errors.Columns.Add("Survey File Value", typeof(double));
                                                dt_errors.Columns.Add("Calculated Value", typeof(string));
                                                dt_errors.Columns.Add("Error Type", typeof(string));

                                                if (dt1.Rows.Count > 0)
                                                {
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        double sta = -1.234;

                                                        if (dt1.Rows[i][Col_Sta3D] != DBNull.Value)
                                                        {
                                                            sta = Convert.ToDouble(dt1.Rows[i][Col_Sta3D]);
                                                        }
                                                        if (sta != -1.234)
                                                        {
                                                            if (_AGEN_mainform.dt_centerline.Rows.Count > 1)
                                                            {
                                                                for (int j = 0; j < _AGEN_mainform.dt_centerline.Rows.Count - 1; ++j)
                                                                {
                                                                    if (_AGEN_mainform.dt_centerline.Rows[j]["3DSta"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["3DSta"] != DBNull.Value &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j]["Y"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j]["Z"] != DBNull.Value &&
                                                                        _AGEN_mainform.dt_centerline.Rows[j + 1]["X"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["Y"] != DBNull.Value && _AGEN_mainform.dt_centerline.Rows[j + 1]["Z"] != DBNull.Value)
                                                                    {
                                                                        double sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j][Col_3DSta]);
                                                                        double sta2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1][Col_3DSta]);

                                                                        if (_AGEN_mainform.dt_centerline.Rows[j][Col_AheadSta] != DBNull.Value)
                                                                        {
                                                                            sta1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j][Col_AheadSta]);
                                                                        }


                                                                        if (_AGEN_mainform.dt_centerline.Rows[j + 1][Col_BackSta] != DBNull.Value)
                                                                        {
                                                                            sta2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1][Col_BackSta]);
                                                                        }

                                                                        if (dt1.Rows[i][Col_x] == DBNull.Value)
                                                                        {
                                                                            if (sta >= sta1 && sta <= sta2)
                                                                            {
                                                                                double x1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["X"]);
                                                                                double y1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Y"]);
                                                                                double z1 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j]["Z"]);
                                                                                double x2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["X"]);
                                                                                double y2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Y"]);
                                                                                double z2 = Convert.ToDouble(_AGEN_mainform.dt_centerline.Rows[j + 1]["Z"]);

                                                                                double x = x1 + (x2 - x1) * (sta - sta1) / (sta2 - sta1);
                                                                                double y = y1 + (y2 - y1) * (sta - sta1) / (sta2 - sta1);
                                                                                double z = z1 + (z2 - z1) * (sta - sta1) / (sta2 - sta1);

                                                                                dt1.Rows[i][Col_x] = x;
                                                                                dt1.Rows[i][Col_y] = y;
                                                                                dt1.Rows[i][Col_z] = z;
                                                                                dt1.Rows[i]["2Dpoint"] = Convert.ToString(x) + "," + Convert.ToString(y);
                                                                                dt1.Rows[i][Col_Sta3D] = sta;

                                                                                dt1.Rows[i][Col_Sta2D] = poly2D.GetDistAtPoint(poly2D.GetClosestPointTo(new Point3d(x, y, poly2D.Elevation), Vector3d.ZAxis, false));

                                                                                j = _AGEN_mainform.dt_centerline.Rows.Count;
                                                                            }

                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        dt_errors.Rows.Add();
                                                                        dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not numeric values omn centerline.xls on row " + (j + _AGEN_mainform.Start_row_CL).ToString();
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        else
                                                        {
                                                            dt_errors.Rows.Add();
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Error Type"] = "Not valid station";
                                                            dt_errors.Rows[dt_errors.Rows.Count - 1]["Calculated Value"] = "See Row " + (start1 + i).ToString();
                                                        }
                                                    }
                                                    transfer_errors_to_panel(dt_errors);
                                                    dt1.Columns.Add("Segment", typeof(string));
                                                    for (int i = 0; i < dt1.Rows.Count; ++i)
                                                    {
                                                        dt1.Rows[i]["Segment"] = _AGEN_mainform.current_segment;
                                                    }
                                                    string nume1 = System.DateTime.Now.Hour + "-" + System.DateTime.Now.Minute;
                                                    Functions.Transfer_datatable_to_new_excel_spreadsheet(dt1, nume1);
                                                }
                                            }
                                        }
                                    }
                                }
                                poly3D.Erase();
                                Trans1.Commit();
                            }
                        }
                        Editor1.SetImpliedSelection(Empty_array);
                        Editor1.WriteMessage("\nCommand:");
                    }
                }
                else
                {
                    MessageBox.Show("no project folder!/r/noperation aborted");
                }
            }
            catch (System.Exception ex)
            {
                _AGEN_mainform.tpage_setup.Set_centerline_label_to_red();
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show(ex.Message);
            }
            set_enable_true();
        }

        private void button_pick_on_profile_Click(object sender, EventArgs e)
        {
            double x0 = 0;
            if (Functions.IsNumeric(textBox_zero_X.Text) == true)
            {
                x0 = Convert.ToDouble(textBox_zero_X.Text);
            }
            double y0 = 0;
            if (Functions.IsNumeric(textBox_zero_Y.Text) == true)
            {
                y0 = Convert.ToDouble(textBox_zero_Y.Text);
            }
            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            List<ObjectId> lista_objid = new List<ObjectId>();
            List<string> lista_chain = new List<string>();

            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;


                        string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();
                        if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                        {
                            ProjFolder = ProjFolder + "\\";
                        }
                        if (System.IO.Directory.Exists(ProjFolder) == true)
                        {
                            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
                            if (System.IO.File.Exists(fisier_cl) == true)
                            {

                                if (_AGEN_mainform.dt_centerline == null)
                                {
                                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                                }

                                System.Data.DataTable dt_cl = _AGEN_mainform.dt_centerline;

                                Polyline poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                                this.MdiParent.WindowState = FormWindowState.Minimized;
                                bool repeat = true;

                                Functions.Creaza_layer(_AGEN_mainform.layer_no_plot, 30, false);
                                Functions.Create_mleader_object_data_table();

                                do
                                {
                                    Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                                    Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                                    PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify point:");
                                    PP1.AllowNone = false;
                                    Point_res1 = Editor1.GetPoint(PP1);

                                    if (Point_res1.Status != PromptStatus.OK)
                                    {

                                        if (lista_chain.Count > 0)
                                        {
                                            Functions.Append_object_data_to_ODXXX(lista_objid, _AGEN_mainform.current_segment, lista_chain);
                                        }

                                        ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                        Trans1.Commit();
                                        this.MdiParent.WindowState = FormWindowState.Normal;
                                        repeat = false;
                                    }

                                    Point3d pt1 = Point_res1.Value;
                                    double dist2d = pt1.X - x0;
                                    double elev = pt1.Y - y0;


                                    if (dist2d >= 0 && dist2d <= poly2D.Length)
                                    {
                                        Point3d pt2d = poly2D.GetPointAtDist(dist2d);
                                        double x = pt2d.X;
                                        double y = pt2d.Y;
                                        double b1 = -1.23456;
                                        double sta = Functions.get_stationCSF_from_point(poly2D, pt2d, dist2d, _AGEN_mainform.dt_centerline, ref b1);
                                        double texth = 2;
                                        if (Functions.IsNumeric(textBox_text_height.Text) == true) texth = Convert.ToDouble(textBox_text_height.Text);

                                        string continut = Functions.Get_chainage_from_double(sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1) + "\r\n" +
                                                            "EL. " + Functions.Get_String_Rounded(elev, 2);

                                        MLeader mlead1 = Functions.creaza_mleader(new Point3d(pt1.X, pt1.Y, 0), continut, texth, texth, texth / 2, texth / 2, texth / 2, 0.1, _AGEN_mainform.layer_no_plot);

                                        lista_objid.Add(mlead1.ObjectId);
                                        lista_chain.Add(Functions.Get_chainage_from_double(sta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1));

                                        Trans1.TransactionManager.QueueForGraphicsFlush();
                                    }








                                } while (repeat == true);
                            }
                        }

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

        private void button_create_offset_Click(object sender, EventArgs e)
        {

            ObjectId[] Empty_array = null;
            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
            Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            if (Functions.IsNumeric(textBox_offset.Text) == false) return;
            double offset1 = Convert.ToDouble(textBox_offset.Text);

            try
            {
                set_enable_false();
                using (DocumentLock lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                        BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        LayerTable LayerTable1 = Trans1.GetObject(ThisDrawing.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                        TextStyleTable Text_style_table1 = Trans1.GetObject(ThisDrawing.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                        Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_centerline;
                        Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_centerline;
                        Prompt_centerline = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the centerline:");
                        Prompt_centerline.SetRejectMessage("\nSelect a polyline!");
                        Prompt_centerline.AllowNone = true;
                        Prompt_centerline.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);
                        Rezultat_centerline = ThisDrawing.Editor.GetEntity(Prompt_centerline);

                        if (Rezultat_centerline.Status != PromptStatus.OK)
                        {
                            Editor1.SetImpliedSelection(Empty_array);
                            Editor1.WriteMessage("\nCommand:");
                            set_enable_true();
                            return;
                        }

                        Polyline poly1 = Trans1.GetObject(Rezultat_centerline.ObjectId, OpenMode.ForRead) as Polyline;
                        if (poly1 != null)
                        {
                            Polyline poly2 = Functions.get_offset_polyline(poly1, offset1);

                            BTrecord.AppendEntity(poly2);
                            Trans1.AddNewlyCreatedDBObject(poly2, true);

                            poly2 = Functions.get_offset_polyline(poly1, -offset1);

                            BTrecord.AppendEntity(poly2);
                            Trans1.AddNewlyCreatedDBObject(poly2, true);

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
