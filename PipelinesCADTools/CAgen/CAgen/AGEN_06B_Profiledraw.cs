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
    public partial class AGEN_Profile_draw : Form
    {
        string Col_Station_ahead = "Station Ahead";
        string Col_Station_back = "Station Back";
        static string Col_3DSta = "3DSta";
        static string Col_BackSta = "BackSta";
        static string Col_AheadSta = "AheadSta";

        bool refresh_attrib_from_blocks = true;

        int rec_no = 0;

        public AGEN_Profile_draw()
        {
            InitializeComponent();

        }

        private void set_enable_false()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_prof_bands);
            lista_butoane.Add(button_insert_label_on_prof);

            lista_butoane.Add(Button_prof_draw);
            lista_butoane.Add(button_show_scan_profile);

            lista_butoane.Add(comboBox_prof_el_lbl_loc);
            lista_butoane.Add(comboBox_prof_textstyle);
            lista_butoane.Add(textBox_overwrite_text_height);
            lista_butoane.Add(textBox_prof_Elev_bottom);
            lista_butoane.Add(textBox_prof_Elev_top);
            lista_butoane.Add(textBox_prof_Hex);
            lista_butoane.Add(textBox_prof_Hspacing);
            lista_butoane.Add(textBox_prof_Vex);
            lista_butoane.Add(textBox_prof_Vspacing);
            lista_butoane.Add(checkBox_draw_ver_at_start);
            lista_butoane.Add(checkBox_elevation);
            lista_butoane.Add(checkBox_cover);


            lista_butoane.Add(checkBox_prof_use_default_grid_val);
            lista_butoane.Add(checkBox_set_zero_at_middle_of_profile);


            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = false;
            }
        }

        private void set_enable_true()
        {
            List<System.Windows.Forms.Control> lista_butoane = new List<Control>();
            lista_butoane.Add(button_draw_prof_bands);
            lista_butoane.Add(button_insert_label_on_prof);

            lista_butoane.Add(Button_prof_draw);

            lista_butoane.Add(button_show_scan_profile);




            lista_butoane.Add(comboBox_prof_el_lbl_loc);
            lista_butoane.Add(comboBox_prof_textstyle);
            lista_butoane.Add(textBox_overwrite_text_height);
            lista_butoane.Add(textBox_prof_Elev_bottom);
            lista_butoane.Add(textBox_prof_Elev_top);
            lista_butoane.Add(textBox_prof_Hex);
            lista_butoane.Add(textBox_prof_Hspacing);
            lista_butoane.Add(textBox_prof_Vex);
            lista_butoane.Add(textBox_prof_Vspacing);
            lista_butoane.Add(checkBox_draw_ver_at_start);
            lista_butoane.Add(checkBox_elevation);
            lista_butoane.Add(checkBox_cover);


            lista_butoane.Add(checkBox_prof_use_default_grid_val);
            lista_butoane.Add(checkBox_set_zero_at_middle_of_profile);

            foreach (System.Windows.Forms.Control bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public string get_comboBox_prof_textstyle()
        {
            return comboBox_prof_textstyle.Text;
        }

        public void set_comboBox_prof_textstyle(string txtname)
        {
            if (comboBox_prof_textstyle.Items.Contains(txtname) == true)
            {
                comboBox_prof_textstyle.SelectedIndex = comboBox_prof_textstyle.Items.IndexOf(txtname);
            }
            else
            {
                comboBox_prof_textstyle.Items.Add(txtname);
                comboBox_prof_textstyle.SelectedIndex = comboBox_prof_textstyle.Items.IndexOf(txtname);
            }

        }

        public int get_comboBox_prof_el_lbl_loc()
        {
            return comboBox_prof_el_lbl_loc.SelectedIndex;
        }

        public bool get_checkBox_draw_ver_at_start()
        {
            return (checkBox_draw_ver_at_start.Checked);
        }

        public void set_checkBox_draw_ver_at_start(bool chck)
        {
            checkBox_draw_ver_at_start.Checked = chck;
        }

        public bool get_checkBox_set_zero_at_middle_of_profile()
        {
            return (checkBox_set_zero_at_middle_of_profile.Checked);
        }

        public void set_checkBox_set_zero_at_middle_of_profile(bool chck)
        {
            checkBox_set_zero_at_middle_of_profile.Checked = chck;
        }

        public bool get_checkBox_hydro_style()
        {
            return (checkBox_hydro_style.Checked);
        }

        public void set_checkBox_hydro_style(bool chck)
        {
            checkBox_hydro_style.Checked = chck;
        }

        public bool get_checkBox_sta_at_90()
        {
            return (checkBox_sta_at_90.Checked);
        }

        public void set_checkBox_sta_at_90(bool chck)
        {
            checkBox_sta_at_90.Checked = chck;
        }





        public int get_textBox_elev_round()
        {
            int nr = 0;
            if (Functions.IsNumeric(textBox_elev_round.Text) == true)
            {
                nr = Convert.ToInt32(textBox_elev_round.Text);
            }
            return nr;
        }

        public void set_comboBox_prof_el_lbl_loc(string txt)
        {
            if (comboBox_prof_el_lbl_loc.Items.Contains(txt) == true)
            {
                comboBox_prof_el_lbl_loc.SelectedIndex = comboBox_prof_el_lbl_loc.Items.IndexOf(txt);
            }

        }




        public string get_textBox_prof_Hex()
        {
            string nr = "1";
            if (Functions.IsNumeric(textBox_prof_Hex.Text) == true)
            {
                nr = textBox_prof_Hex.Text;
            }
            return nr;
        }

        public void set_textBox_prof_Hex(string txt)
        {
            textBox_prof_Hex.Text = txt;
        }


        public string get_textBox_prof_Vex()
        {
            return textBox_prof_Vex.Text;
        }

        public void set_textBox_prof_Vex(string txt)
        {
            textBox_prof_Vex.Text = txt;
        }

        public string get_textBox_prof_Hspacing()
        {
            return textBox_prof_Hspacing.Text;
        }

        public void set_textBox_prof_Hspacing(string txt)
        {
            textBox_prof_Hspacing.Text = txt;
        }

        public string get_textBox_prof_Vspacing()
        {
            return textBox_prof_Vspacing.Text;
        }

        public void set_textBox_prof_Vspacing(string txt)
        {
            textBox_prof_Vspacing.Text = txt;
        }

        public string get_textBox_prof_Elev_top()
        {
            return textBox_prof_Elev_top.Text;
        }

        public string get_textBox_prof_Elev_bottom()
        {
            return textBox_prof_Elev_bottom.Text;
        }

        public void set_textBox_prof_Elev_top(string txt1)
        {
            textBox_prof_Elev_top.Text = txt1;
        }
        public void set_textBox_prof_Elev_bottom(string txt1)
        {
            textBox_prof_Elev_bottom.Text = txt1;
        }


        public bool get_checkBox_prof_use_default_grid_val()
        {

            return checkBox_prof_use_default_grid_val.Checked;
        }



        private void button_profile_refresh_Click(object sender, EventArgs e)
        {

            try
            {
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Functions.Incarca_existing_textstyles_to_combobox(comboBox_prof_textstyle);


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

        private void TextBox_keypress_only_doubles(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_only_pozitive_doubles_at_keypress(sender, e);
        }

        private void TextBox_keypress_elevations(object sender, KeyPressEventArgs e)
        {
            Functions.textbox_input_doubles_at_keypress(sender, e);
        }

        private void checkBox_prof_use_default_grid_val_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_prof_use_default_grid_val.Checked == true) panel_prof_def_values.Visible = false;
            if (checkBox_prof_use_default_grid_val.Checked == false) panel_prof_def_values.Visible = true;
        }

        public void set_checkBox_prof_use_default_grid_val(bool bolval)
        {
            checkBox_prof_use_default_grid_val.Checked = bolval;
        }

        public void set_comboBox_prof_el_lbl_loc(int idx)
        {
            comboBox_prof_el_lbl_loc.SelectedIndex = idx;
        }



        private void button_show_profile_scan_Click(object sender, EventArgs e)
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

            _AGEN_mainform.tpage_profdraw.Hide();
            _AGEN_mainform.tpage_owner_scan.Hide();
            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_profilescan.Show();
        }

        private void button_show_profile_labels_Click(object sender, EventArgs e)
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
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();





        }


        private void Button_graph_prof_draw_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
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

            double Texth = Functions.Get_text_height_from_textstyle(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle());
            if (checkBox_overwrite_text_height.Checked == true && Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
            {
                Texth = Convert.ToDouble(textBox_overwrite_text_height.Text);
            }
            if (Texth == 0)
            {
                MessageBox.Show("the text style you specified does not have a set height\r\nOperation aborted");


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

                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();


                _AGEN_mainform.tpage_profdraw.Show();


                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;

            }


            Ag.WindowState = FormWindowState.Minimized;


            try
            {
                set_enable_false();
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                        PP1.AllowNone = false;
                        Point_res1 = Editor1.GetPoint(PP1);

                        if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                        {
                            Ag.WindowState = FormWindowState.Normal;
                            set_enable_true();
                            ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                            return;
                        }






                        double Hexag = 1;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
                        {
                            Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
                        }

                        double Vexag = 1;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex()) == true)
                        {
                            Vexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex());
                        }

                        double Hincr = 100;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hspacing()) == true)
                        {
                            Hincr = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hspacing());
                        }

                        double vincr = 100;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vspacing()) == true)
                        {
                            vincr = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vspacing());
                        }

                        string fisier_prof = "";

                        string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();



                        if (System.IO.Directory.Exists(ProjFolder) == true)
                        {
                            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                            {
                                ProjFolder = ProjFolder + "\\";
                            }
                            fisier_prof = ProjFolder + _AGEN_mainform.prof_excel_name;
                            if (System.IO.File.Exists(fisier_prof) == false)
                            {
                                MessageBox.Show("the profile data file does not exist");
                                set_enable_true();
                                Ag.WindowState = FormWindowState.Normal;
                                return;
                            }

                            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

                            if (System.IO.File.Exists(fisier_cl) == true) //&& _AGEN_mainform.tpage_sheetindex.get_checkBox_station_equations_value() == true
                            {
                                _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                            }
                        }
                        else
                        {
                            MessageBox.Show("the project database folder does not exist");
                            set_enable_true();
                            Ag.WindowState = FormWindowState.Normal;
                            return;
                        }
                        System.Data.DataTable dt_top = new System.Data.DataTable();
                        System.Data.DataTable dt_prof = Load_existing_profile_graph(fisier_prof, ref dt_top);


                        bool L1 = true;
                        if (_AGEN_mainform.tpage_profdraw.get_comboBox_prof_el_lbl_loc() == 2)
                        {
                            L1 = false;
                        }

                        bool L2 = true;
                        if (_AGEN_mainform.tpage_profdraw.get_comboBox_prof_el_lbl_loc() == 1)
                        {
                            L2 = false;
                        }

                        string Suff = "'";
                        if (_AGEN_mainform.units_of_measurement == "m")
                        {
                            Suff = "";
                        }


                        if (System.IO.Directory.Exists(ProjFolder) == true && dt_prof.Rows.Count > 0)
                        {
                            if (System.IO.File.Exists(fisier_prof) == true)
                            {

                                double Min_el = 100000;
                                double Max_el = -100000;

                                for (int i = 0; i < dt_prof.Rows.Count; ++i)
                                {
                                    if (dt_prof.Rows[i][_AGEN_mainform.Col_Elev] != DBNull.Value)
                                    {
                                        double z1 = Convert.ToDouble(dt_prof.Rows[i][_AGEN_mainform.Col_Elev]);
                                        if (z1 > Max_el) Max_el = z1;
                                        if (z1 < Min_el) Min_el = z1;
                                    }
                                }

                                double Downelev = Functions.Round_Down_as_double(Min_el, vincr) - 10 * vincr;
                                double Upelev = Functions.Round_Up_as_double(Max_el, vincr) + 10 * vincr;
                                if (_AGEN_mainform.tpage_profdraw.get_checkBox_prof_use_default_grid_val() == false)
                                {
                                    string Del_s = _AGEN_mainform.tpage_profdraw.get_textBox_prof_Elev_bottom();
                                    string Uel_s = _AGEN_mainform.tpage_profdraw.get_textBox_prof_Elev_top();
                                    if (Functions.IsNumeric(Del_s) == true)
                                    {
                                        Downelev = Functions.Round_Down_as_double(Convert.ToDouble(Del_s), vincr);
                                    }
                                    if (Functions.IsNumeric(Uel_s) == true)
                                    {
                                        Upelev = Functions.Round_Up_as_double(Convert.ToDouble(Uel_s), vincr);
                                    }
                                }

                                _AGEN_mainform.tpage_profdraw.set_textBox_prof_Elev_top(Upelev.ToString());
                                _AGEN_mainform.tpage_profdraw.set_textBox_prof_Elev_bottom(Downelev.ToString());

                                Functions.create_backup(_AGEN_mainform.config_path);
                                _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);


                                if (_AGEN_mainform.dt_station_equation != null)
                                {
                                    if (_AGEN_mainform.dt_station_equation.Rows.Count > 0 && _AGEN_mainform.COUNTRY == "USA")
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


                                                Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                                                double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                                double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                                _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                            }
                                        }
                                    }
                                }

                                int elev_round = 0;
                                if (Functions.IsNumeric(textBox_elev_round.Text) == true)
                                {
                                    elev_round = Math.Abs(Convert.ToInt32(textBox_elev_round.Text));
                                }

                                bool draw_pipe = false;

                                textBox_overwrite_text_height.Text = Convert.ToString(Texth);
                                checkBox_overwrite_text_height.Checked = true;

                                Functions.Draw_grid_profile(dt_prof, dt_top, Point_res1.Value, Hincr, vincr, Hexag, Vexag, Downelev, Upelev, elev_round,
                                                            _AGEN_mainform.layer_prof_grid, _AGEN_mainform.layer_prof_text, _AGEN_mainform.layer_prof_ground,
                                                            _AGEN_mainform.layer_prof_pipe, Texth,
                                                                    Functions.Get_textstyle_id(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle()),
                                                                            Suff, L1, L2, _AGEN_mainform.config_path, _AGEN_mainform.ExcelVisible, _AGEN_mainform.Start_row_1,
                                                                                 _AGEN_mainform.units_of_measurement, _AGEN_mainform.dt_station_equation, draw_pipe);

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
            set_enable_true();



            _AGEN_mainform.tpage_setup.button_align_config_saveall_boolean(true);



            Ag.WindowState = FormWindowState.Normal;

        }
        public bool get_checkbox_offset()
        {
            return checkBox_offset_top.Checked;
        }

        public double get_textbox_offset()
        {
            string val1 = textBox_offset.Text;
            if (Functions.IsNumeric(val1) == true)
            {
                return Convert.ToDouble(val1);
            }
            else
            {
                return 0;
            }

        }

        public void set_checkBox_overwrite_text_height(bool val1 = false)
        {
            checkBox_overwrite_text_height.Checked = val1;
        }

        public void set_textBox_overwrite_text_height(string val1 = "5")
        {
            textBox_overwrite_text_height.Text = val1;
        }

        public bool get_checkBox_overwrite_text_height()
        {
            return checkBox_overwrite_text_height.Checked;
        }

        public string get_textBox_overwrite_text_height()
        {
            if (Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
            {
                return textBox_overwrite_text_height.Text;
            }
            else
            {
                return "";
            }
        }

        public System.Data.DataTable Load_existing_profile_graph(string profxl, ref System.Data.DataTable dt_top)
        {
            //if dt_top == null then no top load

            if (System.IO.File.Exists(profxl) == false)
            {
                MessageBox.Show("the profile data file does not exist");
                return null;
            }


            System.Data.DataTable dt2 = new System.Data.DataTable();


            Microsoft.Office.Interop.Excel.Application Excel1 = null;
            Microsoft.Office.Interop.Excel.Worksheet W_prof = null;
            Microsoft.Office.Interop.Excel.Workbook Workboook_prof = null;
            Microsoft.Office.Interop.Excel.Worksheet W_top = null;

            bool is_opened_prof = false;
            bool close_prof = false;
            try
            {
                try
                {
                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                    {
                        string workbookname = Workbook2.FullName;

                        if (workbookname.ToLower() == profxl.ToLower())
                        {
                            Workboook_prof = Workbook2;
                            is_opened_prof = true;
                        }

                    }
                }
                catch (System.Exception ex)
                {
                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                }

                if (is_opened_prof == false)
                {
                    Workboook_prof = Excel1.Workbooks.Open(profxl);
                    close_prof = true;
                }

                W_prof = Workboook_prof.Worksheets[1];



                if (dt_top != null)
                {
                    foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workboook_prof.Worksheets)
                    {
                        if (Wx.Name == "TOP")
                        {
                            W_top = Wx;

                        }
                    }
                }

                dt2 = Functions.Build_Data_table_profile_from_excel(W_prof, _AGEN_mainform.Start_row_graph_profile + 1);

                if (W_top != null)
                {
                    dt_top = Functions.Build_Data_table_profile_from_excel(W_top, _AGEN_mainform.Start_row_graph_profile + 1);
                }



                if (close_prof == true)
                {
                    Workboook_prof.Close();
                }


                if (Excel1.Workbooks.Count == 0) Excel1.Quit();


            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (W_prof != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_prof);
                if (W_top != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_top);
                if (Workboook_prof != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workboook_prof);
                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
            }

            return dt2;

        }

        private void label_draw_profile_Click(object sender, EventArgs e)
        {


            if (panel_dan.Visible == false)
            {
                panel_dan.Visible = true;

            }
            else
            {
                panel_dan.Visible = false;


            }

        }

        public bool get_checkbox_pipes_value()
        {
            return checkBox_pipes.Checked;
        }

        #region profile label
        private void button_insert_labels_on_profile_Click(object sender, EventArgs e)
        {
            Functions.Kill_excel();

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
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

            double Texth = Functions.Get_text_height_from_textstyle(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle());

            if (checkBox_overwrite_text_height.Checked == true && Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
            {
                Texth = Convert.ToDouble(textBox_overwrite_text_height.Text);
            }
            if (Texth <= 0) Texth = 10;

            if (Ag != null)
            {


                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
            }

            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();


            if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
            {
                ProjFolder = ProjFolder + "\\";
            }

            string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                MessageBox.Show("No centerline file found");
                return;
            }

            string fisier_cs = ProjFolder + _AGEN_mainform.crossing_excel_name;

            if (System.IO.File.Exists(fisier_cs) == false)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the crossing data file does not exist");
                return;
            }


            double Hexag = 0;
            if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
            {
                Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("specify the profile horizontal exxageration");
                return;
            }

            _AGEN_mainform.tpage_processing.Show();
            _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);

            bool defineste_block = false;
            if (checkBox_pipes.Checked == true && (checkBox_cover.Checked == true || checkBox_elevation.Checked == true))
            {
                defineste_block = true;
            }

            _AGEN_mainform.Data_Table_crossings = _AGEN_mainform.tpage_crossing_draw.Load_existing_crossing(fisier_cs, "", defineste_block);

            string fisier_prof = ProjFolder + _AGEN_mainform.prof_excel_name;
            if (System.IO.File.Exists(fisier_prof) == false)
            {
                MessageBox.Show("the profile data file does not exist");
                set_enable_true();
                Ag.WindowState = FormWindowState.Normal;
                return;
            }

            System.Data.DataTable dt_null = null;

            System.Data.DataTable dt_profile = Load_existing_profile_graph(fisier_prof, ref dt_null);


            if (_AGEN_mainform.Data_Table_crossings != null)
            {
                if (_AGEN_mainform.Data_Table_crossings.Rows.Count > 0)
                {

                    ObjectId[] Empty_array = null;
                    Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                    Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                    Matrix3d curent_ucs_matrix = Editor1.CurrentUserCoordinateSystem;
                    Editor1.SetImpliedSelection(Empty_array);
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();


                    try
                    {
                        set_enable_false();
                        using (DocumentLock lock1 = ThisDrawing.LockDocument())
                        {
                            using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                            {

                                _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                                _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);

                                int lr = 1;

                                if (_AGEN_mainform.Left_to_Right == false)
                                {
                                    lr = -1;
                                }

                                BlockTable BlockTable1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                                double ymin = -1000000;
                                double ymax = 1000000;

                                Polyline Poly2d = new Polyline();


                                bool exista1 = false;

                                for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                {


                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][13] != DBNull.Value)
                                    {
                                        string val1 = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][13]);
                                        if (val1.ToLower() == "yes" || val1.ToLower() == "y")
                                        {
                                            exista1 = true;
                                            i = _AGEN_mainform.Data_Table_crossings.Rows.Count;
                                        }

                                    }
                                }

                                if (exista1 == false)
                                {
                                    _AGEN_mainform.tpage_processing.Hide();
                                    set_enable_true();
                                    MessageBox.Show("the crossing table column DispProf does not have any YES\r\noperation aborted");
                                    return;
                                }

                                List<ObjectId> lista_poly = new List<ObjectId>();
                                List<double> lista_start = new List<double>();
                                List<double> lista_end = new List<double>();

                                Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

                                string Agen_profile_band_V2 = "Agen_profile_band_V2";
                                string Agen_profile_band_V3 = "Agen_profile_band_V3";

                                if (Tables1.IsTableDefined(Agen_profile_band_V2) == true)
                                {
                                    foreach (ObjectId id1 in BTrecord)
                                    {
                                        Polyline poly_ground = Trans1.GetObject(id1, OpenMode.ForRead) as Polyline;
                                        if (poly_ground != null)
                                        {

                                            if (Tables1.IsTableDefined(Agen_profile_band_V2) == true)
                                            {
                                                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V2])
                                                {

                                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                    {
                                                        if (Records1.Count > 0)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                double start1 = -123.4;
                                                                double end1 = -123.4;
                                                                string segm1 = "123456";
                                                                for (int i = 0; i < Record1.Count; ++i)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                    string Nume_field = Field_def1.Name;
                                                                    string Valoare_field = Record1[i].StrValue;

                                                                    if (Nume_field.ToLower() == "beginsta")
                                                                    {
                                                                        if (Functions.IsNumeric(Valoare_field) == true)
                                                                        {
                                                                            start1 = Convert.ToDouble(Valoare_field);
                                                                        }
                                                                    }

                                                                    if (Nume_field.ToLower() == "endsta")
                                                                    {
                                                                        if (Functions.IsNumeric(Valoare_field) == true)
                                                                        {
                                                                            end1 = Convert.ToDouble(Valoare_field);
                                                                        }
                                                                    }
                                                                    if (Nume_field.ToLower() == "segment")
                                                                    {
                                                                        segm1 = Convert.ToString(Valoare_field);
                                                                    }
                                                                }

                                                                string segment2 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                                if (_AGEN_mainform.tpage_setup.Get_segment_name1() == "not defined")
                                                                {
                                                                    segment2 = "";
                                                                }

                                                                if (start1 != -123.4 && end1 != 123.4 && segm1.ToLower() == segment2.ToLower())
                                                                {
                                                                    lista_poly.Add(id1);
                                                                    lista_start.Add(start1);
                                                                    lista_end.Add(end1);
                                                                }

                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                            if (Tables1.IsTableDefined(Agen_profile_band_V3) == true)
                                            {
                                                using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V3])
                                                {

                                                    using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), id1, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                    {
                                                        if (Records1.Count > 0)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                            foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                            {
                                                                double start1 = -123.4;
                                                                double end1 = -123.4;
                                                                string segm1 = "123456";
                                                                for (int i = 0; i < Record1.Count; ++i)
                                                                {
                                                                    Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[i];
                                                                    string Nume_field = Field_def1.Name;
                                                                    string Valoare_field = Record1[i].StrValue;

                                                                    if (Nume_field.ToLower() == "beginsta")
                                                                    {
                                                                        if (Functions.IsNumeric(Valoare_field) == true)
                                                                        {
                                                                            start1 = Convert.ToDouble(Valoare_field);
                                                                        }
                                                                    }

                                                                    if (Nume_field.ToLower() == "endsta")
                                                                    {
                                                                        if (Functions.IsNumeric(Valoare_field) == true)
                                                                        {
                                                                            end1 = Convert.ToDouble(Valoare_field);
                                                                        }
                                                                    }
                                                                    if (Nume_field.ToLower() == "segment")
                                                                    {
                                                                        segm1 = Convert.ToString(Valoare_field);
                                                                    }
                                                                }

                                                                string segment2 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                                if (_AGEN_mainform.tpage_setup.Get_segment_name1() == "not defined")
                                                                {
                                                                    segment2 = "";
                                                                }

                                                                if (start1 != -123.4 && end1 != 123.4 && segm1.ToLower() == segment2.ToLower())
                                                                {
                                                                    lista_poly.Add(id1);
                                                                    lista_start.Add(start1);
                                                                    lista_end.Add(end1);
                                                                }

                                                            }
                                                        }
                                                    }

                                                }
                                            }

                                        }
                                    }
                                }




                                Functions.Creaza_layer(_AGEN_mainform.layer_prof_block_labels, 2, true);

                                if ((checkBox_cover.Checked == true || checkBox_elevation.Checked == true) && checkBox_pipes.Checked == true)
                                {
                                    Functions.Creaza_layer("Agen_symbols", 2, true);
                                    Functions.Creaza_layer("NO PLOT", 40, false);
                                }

                                ObjectId text_style_id = Functions.Get_textstyle_id(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle());

                                if (Texth <= 0) Texth = 10;


                                for (int i = 0; i < _AGEN_mainform.Data_Table_crossings.Rows.Count; ++i)
                                {
                                    bool creaza_block = false;
                                    bool creaza_mleader = false;
                                    string block_name = "";
                                    string at_sta = "";
                                    string at_desc = "";
                                    double z = 0;
                                    double z_on_cl = 0;

                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][13] != DBNull.Value)
                                    {
                                        string val1 = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][13]);
                                        if (val1.ToLower() == "yes" || val1.ToLower() == "y" || val1.ToLower() == "true")
                                        {
                                            if (_AGEN_mainform.Data_Table_crossings.Rows[i][15] != DBNull.Value)
                                            {
                                                block_name = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][15]);
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][16] != DBNull.Value)
                                                {
                                                    at_sta = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][16]);
                                                }
                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][17] != DBNull.Value)
                                                {
                                                    at_desc = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][17]);
                                                }
                                                creaza_block = true;
                                            }
                                            else
                                            {
                                                creaza_mleader = true;
                                            }
                                        }
                                    }

                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_Elev] != DBNull.Value)
                                    {
                                        z = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_Elev]);
                                    }




                                    if (creaza_block == true || creaza_mleader == true)
                                    {
                                        double Station = -1;
                                        double Station_2d = -1;
                                        Point3d pt_on_2d = new Point3d();

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][1] != DBNull.Value)
                                        {
                                            Station = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][1]);
                                        }

                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][2] != DBNull.Value)
                                        {
                                            Station = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][2]);
                                        }



                                        if (dt_profile.Rows.Count > 1)
                                        {
                                            for (int j = 0; j < dt_profile.Rows.Count - 1; ++j)
                                            {
                                                if (dt_profile.Rows[j]["Station"] != DBNull.Value && dt_profile.Rows[j]["Elev"] != DBNull.Value &&
                                                    dt_profile.Rows[j + 1]["Station"] != DBNull.Value && dt_profile.Rows[j + 1]["Elev"] != DBNull.Value)
                                                {
                                                    double sta1 = Convert.ToDouble(dt_profile.Rows[j]["Station"]);
                                                    double sta2 = Convert.ToDouble(dt_profile.Rows[j + 1]["Station"]);
                                                    double elev1 = Convert.ToDouble(dt_profile.Rows[j]["Elev"]);
                                                    double elev2 = Convert.ToDouble(dt_profile.Rows[j + 1]["Elev"]);
                                                    if (Station >= sta1 && Station <= sta2)
                                                    {
                                                        z_on_cl = elev1 + (Station - sta1) * (elev2 - elev1) / (sta2 - sta1);
                                                        j = dt_profile.Rows.Count;
                                                    }
                                                }
                                            }
                                        }


                                        if (_AGEN_mainform.COUNTRY == "CANADA" &&
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x] != DBNull.Value &&
                                            _AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y] != DBNull.Value)
                                        {
                                            double x = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_x]);
                                            double y = Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i][_AGEN_mainform.Col_y]);
                                            pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, _AGEN_mainform.Poly2D.Elevation), Vector3d.ZAxis, false);
                                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                                            Station = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                                            Station_2d = _AGEN_mainform.Poly2D.GetDistanceAtParameter(param1);
                                            double b1 = -1.23456;
                                            Station = Functions.get_stationCSF_from_point(_AGEN_mainform.Poly2D, pt_on_2d, Station_2d, _AGEN_mainform.dt_centerline, ref b1);
                                        }


                                        double vexag = 0;
                                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex()) == true)
                                        {
                                            vexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex());
                                        }
                                        else
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("specify the profile vertical exaggeration");
                                            return;
                                        }

                                        if (lista_start.Count > 0 && lista_start.Count == lista_end.Count && lista_start.Count == lista_poly.Count)
                                        {
                                            for (int k = 0; k < lista_poly.Count; ++k)
                                            {
                                                if (lista_poly[k] != null && lista_poly[k] != ObjectId.Null)
                                                {
                                                    Poly2d = Trans1.GetObject(lista_poly[k], OpenMode.ForRead) as Polyline;
                                                    if (Poly2d != null)
                                                    {
                                                        double start1 = lista_start[k];
                                                        double end1 = lista_end[k];
                                                        if (Station >= start1 && Station <= end1)
                                                        {
                                                            for (int n = 0; n < Poly2d.NumberOfVertices - 1; ++n)
                                                            {
                                                                double y = Poly2d.GetPointAtParameter(n).Y;
                                                                if (n == 0)
                                                                {
                                                                    ymin = y;
                                                                    ymax = y;
                                                                }
                                                                else
                                                                {
                                                                    if (ymin > y)
                                                                    {
                                                                        ymin = y;
                                                                    }
                                                                    if (ymax < y)
                                                                    {
                                                                        ymax = y;
                                                                    }
                                                                }
                                                            }

                                                            double x1 = Poly2d.StartPoint.X + lr * (Station - start1) * Hexag;


                                                            if (_AGEN_mainform.COUNTRY == "CANADA" && _AGEN_mainform.dt_station_equation != null && _AGEN_mainform.dt_station_equation.Rows.Count > 0)
                                                            {


                                                                double ahead0 = start1;
                                                                double dif1 = 0;

                                                                for (int j = 0; j < _AGEN_mainform.dt_station_equation.Rows.Count; ++j)
                                                                {
                                                                    if (_AGEN_mainform.dt_station_equation.Rows[j][Col_Station_ahead] != DBNull.Value && _AGEN_mainform.dt_station_equation.Rows[j][Col_Station_back] != DBNull.Value)
                                                                    {
                                                                        double back1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[j][Col_Station_back]);
                                                                        double ahead1 = Convert.ToDouble(_AGEN_mainform.dt_station_equation.Rows[j][Col_Station_ahead]);

                                                                        if (start1 <= back1 && ahead1 <= end1)
                                                                        {
                                                                            if (Station > ahead1)
                                                                            {
                                                                                dif1 = dif1 + back1 - ahead0;
                                                                                ahead0 = ahead1;
                                                                            }

                                                                        }
                                                                    }
                                                                }



                                                                x1 = Poly2d.StartPoint.X + lr * (dif1 + (Station - ahead0)) * Hexag;

                                                            }

                                                            Line line1 = new Line(new Point3d(x1, ymin - 10000, Poly2d.Elevation), new Point3d(x1, ymax + 10000, Poly2d.Elevation));

                                                            Point3dCollection col1 = Functions.Intersect_on_both_operands(Poly2d, line1);


                                                            if (col1.Count == 0)
                                                            {
                                                                col1.Add(new Point3d(x1, Poly2d.GetPoint2dAt(0).Y, Poly2d.Elevation));
                                                            }




                                                            for (int n = 0; n < col1.Count; ++n)
                                                            {
                                                                Point3d inspt = new Point3d();

                                                                inspt = col1[n];

                                                                if (checkBox_cover.Checked == true)
                                                                {
                                                                    inspt = new Point3d(inspt.X, inspt.Y - z * vexag, inspt.Z);
                                                                }
                                                                else if (checkBox_elevation.Checked == true)
                                                                {
                                                                    inspt = new Point3d(inspt.X, inspt.Y - (z_on_cl - z) * vexag, inspt.Z);
                                                                }


                                                                string descriptie = "";
                                                                if (_AGEN_mainform.Data_Table_crossings.Rows[i][6] != DBNull.Value)
                                                                {
                                                                    descriptie = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][6]);
                                                                }



                                                                double dispsta = Station;

                                                                if (_AGEN_mainform.COUNTRY == "USA")
                                                                {
                                                                    dispsta = Functions.Station_equation_of(Station, _AGEN_mainform.dt_station_equation);
                                                                }

                                                                string display_sta_string = Functions.Get_chainage_from_double(dispsta, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);


                                                                string mleader_descr = display_sta_string + " " + descriptie;

                                                                if (checkBox_no_station.Checked == true)
                                                                {
                                                                    mleader_descr = descriptie;
                                                                }


                                                                System.Collections.Specialized.StringCollection col_atr = new System.Collections.Specialized.StringCollection();
                                                                System.Collections.Specialized.StringCollection col_val = new System.Collections.Specialized.StringCollection();

                                                                col_atr.Add(at_desc + "1");
                                                                col_val.Add(descriptie);

                                                                col_atr.Add(at_sta + "1");
                                                                col_val.Add(display_sta_string);

                                                                col_atr.Add(at_desc);
                                                                col_val.Add(descriptie);

                                                                if (checkBox_no_station.Checked == false)
                                                                {
                                                                    col_atr.Add(at_sta);
                                                                    col_val.Add(display_sta_string);
                                                                }


                                                                if (creaza_block == true && defineste_block == false)
                                                                {
                                                                    BlockReference br1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "",
                                                                              block_name, inspt, 1 / _AGEN_mainform.Vw_scale, 0, _AGEN_mainform.layer_prof_block_labels, col_atr, col_val);

                                                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Visibility"] != DBNull.Value)
                                                                    {
                                                                        string vis = "xxx";
                                                                        vis = Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i]["Visibility"]);
                                                                        Functions.set_block_visibility(br1, vis);
                                                                    }


                                                                }
                                                                else if (creaza_mleader == true && defineste_block == false)
                                                                {
                                                                    Functions.Create_mleader_on_profile_with_database(ThisDrawing.Database, BTrecord, inspt, _AGEN_mainform.layer_prof_block_labels, mleader_descr, Texth, text_style_id);
                                                                }
                                                                else if (defineste_block == true)
                                                                {
                                                                    if (_AGEN_mainform.Data_Table_crossings.Rows[i]["Pipe Size in feet"] != DBNull.Value)
                                                                    {
                                                                        #region block creation

                                                                        double diam1 = Math.Abs(Convert.ToDouble(_AGEN_mainform.Data_Table_crossings.Rows[i]["Pipe Size in feet"]));
                                                                        string name_of_block = "_" + diam1;

                                                                        string content1 = display_sta_string;
                                                                        if (_AGEN_mainform.Data_Table_crossings.Rows[i][6] != DBNull.Value)
                                                                        {
                                                                            content1 = display_sta_string + "\r\n" + Convert.ToString(_AGEN_mainform.Data_Table_crossings.Rows[i][6]);
                                                                        }


                                                                        if (BlockTable1.Has(name_of_block) == false)
                                                                        {
                                                                            BlockTable1.UpgradeOpen();
                                                                            using (BlockTableRecord bltrec1 = new BlockTableRecord())
                                                                            {
                                                                                bltrec1.Name = name_of_block;


                                                                                Circle cerc1 = new Circle(new Point3d(0, -diam1 / 2, 0), Vector3d.ZAxis, diam1 / 2);
                                                                                cerc1.Layer = "0";
                                                                                bltrec1.AppendEntity(cerc1);



                                                                                AttributeDefinition att1 = new AttributeDefinition();
                                                                                att1.Tag = "DESCRIPTION";
                                                                                att1.Layer = "NO PLOT";
                                                                                att1.Height = diam1 / 4;
                                                                                att1.Position = new Point3d(0, -diam1 - diam1 / 4, 0);
                                                                                att1.IsMTextAttributeDefinition = true;
                                                                                att1.Justify = AttachmentPoint.TopCenter;

                                                                                att1.TextString = content1;
                                                                                bltrec1.AppendEntity(att1);



                                                                                BlockTable1.Add(bltrec1);
                                                                                Trans1.AddNewlyCreatedDBObject(bltrec1, true);

                                                                                col_atr = new System.Collections.Specialized.StringCollection();
                                                                                col_val = new System.Collections.Specialized.StringCollection();
                                                                                col_atr.Add("DESCRIPTION");
                                                                                col_val.Add(content1);

                                                                                BlockReference b1 = Functions.InsertBlock_with_multiple_atributes_with_database_2_SCALES(ThisDrawing.Database, BTrecord, "", name_of_block, inspt, Hexag, vexag, 0, "Agen_symbols", col_atr, col_val);
                                                                                b1.ColorIndex = 256;
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            col_atr = new System.Collections.Specialized.StringCollection();
                                                                            col_val = new System.Collections.Specialized.StringCollection();
                                                                            col_atr.Add("DESCRIPTION");
                                                                            col_val.Add(content1);

                                                                            BlockReference b1 = Functions.InsertBlock_with_multiple_atributes_with_database_2_SCALES(ThisDrawing.Database, BTrecord, "", name_of_block, inspt, Hexag, vexag, 0, "Agen_symbols", col_atr, col_val);
                                                                            b1.ColorIndex = 256;
                                                                        }

                                                                        #endregion
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
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
                    Editor1.WriteMessage("\nCommand:");

                }
                else
                {
                    MessageBox.Show("no crossing data found");
                }
            }
            else
            {
                MessageBox.Show("no crossing data found");
            }
            _AGEN_mainform.tpage_processing.Hide();
            set_enable_true();
        }






        private string magic_number_wrap(string string1, int magic_no)
        {
            string new_string = "";
            string string_de_procesat = string1;
            do
            {
                string_de_procesat = string_de_procesat.Replace("  ", " ");
            } while (string_de_procesat.Contains("  ") == true);

            if (string_de_procesat.Substring(0, 1) == " ")
            {
                string_de_procesat = string_de_procesat.Substring(1);
            }

            if (string1.Contains(" ") == true)
            {
                string[] cuvinte; ;
                char spatiu = Convert.ToChar(" ");
                cuvinte = string_de_procesat.Split(spatiu);
                for (int i = 0; i < cuvinte.Length; ++i)
                {
                    string cuvant1 = cuvinte[i];
                    if (cuvant1.Length < magic_no)
                    {
                        if (i + 1 < cuvinte.Length)
                        {
                            int len = cuvant1.Length;
                            do
                            {
                                for (int j = i + 1; j < cuvinte.Length; ++j)
                                {
                                    string cuvant2 = cuvinte[j];
                                    if (cuvant1.Length + 1 + cuvant2.Length < magic_no)
                                    {
                                        cuvant1 = cuvant1 + " " + cuvant2;
                                        len = cuvant1.Length;
                                        i = i + 1;
                                    }
                                    else
                                    {
                                        new_string = return_new_string(new_string, cuvant1);
                                        len = magic_no;
                                        j = cuvinte.Length;
                                    }
                                    if (j == cuvinte.Length - 1)
                                    {
                                        new_string = return_new_string(new_string, cuvant1);
                                        len = magic_no;
                                    }
                                }
                            } while (len < magic_no);
                        }
                        else
                        {
                            new_string = return_new_string(new_string, cuvant1);
                        }
                    }
                    else
                    {
                        new_string = return_new_string(new_string, cuvant1);
                    }
                }
            }
            else
            {
                new_string = string1;
            }
            return new_string;
        }
        private string return_new_string(string new_string, string cuvant1)
        {
            if (new_string != "")
            {
                new_string = new_string + "\r\n" + cuvant1;
            }
            else
            {
                new_string = cuvant1;
            }

            return new_string;
        }


        #endregion

        public void button_load_data_for_profile_band_Click(object sender, EventArgs e)
        {
            set_enable_false();
            Functions.Kill_excel();


            try
            {
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    set_enable_true();
                    return;
                }
                string fisier_prof_band = "";
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                    {
                        ProjFolder = ProjFolder + "\\";
                    }

                    fisier_prof_band = ProjFolder + _AGEN_mainform.band_prof_excel_name;
                    if (System.IO.File.Exists(fisier_prof_band) == false)
                    {
                        MessageBox.Show("the profile band data file does not exist");
                        set_enable_true();
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("the project database folder does not exist");
                    set_enable_true();
                    return;
                }

                _AGEN_mainform.Data_Table_profile_band = Load_existing_profile_band_data(fisier_prof_band);

                set_enable_true();
            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            set_enable_true();
        }

        public System.Data.DataTable Load_existing_profile_band_data(string File1)
        {

            if (System.IO.File.Exists(File1) == false)
            {
                MessageBox.Show("the profile band data file does not exist");
                return null;
            }


            System.Data.DataTable dt1 = new System.Data.DataTable();

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
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(File1);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];

                try
                {
                    dt1 = Functions.Build_Data_table_profile_band_from_excel(W1, _AGEN_mainform.Start_row_profile_band + 1);
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
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }
            return dt1;

        }

        private void generate_profile_band_file_from_sheet_index()
        {
            Functions.Kill_excel();

            try
            {
                _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;

                if (System.IO.File.Exists(_AGEN_mainform.config_path) == false)
                {
                    MessageBox.Show("no config file loaded\r\nOperation aborted");
                    return;
                }
                string fisier_prof_band = "";
                string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

                if (System.IO.Directory.Exists(ProjFolder) == true)
                {
                    if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                    {
                        ProjFolder = ProjFolder + "\\";
                    }
                    fisier_prof_band = ProjFolder + _AGEN_mainform.band_prof_excel_name;
                    if (System.IO.File.Exists(fisier_prof_band) == false)

                    {
                        string fisier_si = ProjFolder + _AGEN_mainform.sheet_index_excel_name;
                        if (System.IO.File.Exists(fisier_si) == true)
                        {
                            _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);

                            _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                            _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

                           
                            if (_AGEN_mainform.Data_Table_profile_band == null || _AGEN_mainform.Data_Table_profile_band.Rows.Count == 0)
                            {
                                _AGEN_mainform.Data_Table_profile_band = Functions.Creaza_profile_band_datatable_structure();

                                for (int i = 0; i < _AGEN_mainform.dt_sheet_index.Rows.Count; ++i)
                                {
                                    double M1 = -1;
                                    double M2 = -1;
                                    if (_AGEN_mainform.dt_sheet_index.Rows[i]["StaBeg"] != DBNull.Value &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["StaBeg"])) == true &&
                                        _AGEN_mainform.dt_sheet_index.Rows[i]["StaEnd"] != DBNull.Value &&
                                        Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["StaEnd"])) == true &&
                                        _AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"] != DBNull.Value)
                                    {

                                        M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["StaBeg"]);
                                        M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[i]["StaEnd"]);


                                        string dwg = Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[i]["DwgNo"]);
                                        _AGEN_mainform.Data_Table_profile_band.Rows.Add();
                                        _AGEN_mainform.Data_Table_profile_band.Rows[_AGEN_mainform.Data_Table_profile_band.Rows.Count - 1]["DwgNo"] = dwg;
                                        _AGEN_mainform.Data_Table_profile_band.Rows[_AGEN_mainform.Data_Table_profile_band.Rows.Count - 1]["StaBeg"] = M1;
                                        _AGEN_mainform.Data_Table_profile_band.Rows[_AGEN_mainform.Data_Table_profile_band.Rows.Count - 1]["StaEnd"] = M2;

                                        double zero = 0;

                                        if (checkBox_set_zero_at_middle_of_profile.Checked == true)
                                        {
                                            zero = (M1 + M2) / 2;
                                        }

                                        _AGEN_mainform.Data_Table_profile_band.Rows[_AGEN_mainform.Data_Table_profile_band.Rows.Count - 1]["Zero_position"] = zero;

                                    }

                                }
                            }
                        }
                    }
                    else
                    {
                        _AGEN_mainform.Data_Table_profile_band = Load_existing_profile_band_data(fisier_prof_band);

                    }

                }
                else
                {
                    MessageBox.Show("the project database folder does not exist");
                    return;
                }



            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }



        private void button_draw_prof_bands_Click(object sender, EventArgs e)
        {

            Functions.Kill_excel();

            if (_AGEN_mainform.Vw_profband_height == 0)
            {
                MessageBox.Show("Profile band height = 0, verify your viewport settings!");
                set_enable_true();
                return;
            }

            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
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

            if (Functions.Get_if_workbook_is_open_in_Excel(_AGEN_mainform.band_prof_excel_name) == true)
            {
                MessageBox.Show("Please close the " + _AGEN_mainform.band_prof_excel_name + " file");
                return;
            }




            double Texth = Functions.Get_text_height_from_textstyle(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle());
            if (checkBox_overwrite_text_height.Checked == true && Functions.IsNumeric(textBox_overwrite_text_height.Text) == true)
            {
                Texth = Convert.ToDouble(textBox_overwrite_text_height.Text);
            }
            if (Texth <= 0) Texth = 10;
            set_enable_false();




            if (_AGEN_mainform.Data_Table_profile_band == null || _AGEN_mainform.Data_Table_profile_band.Rows.Count == 0)
            {
                generate_profile_band_file_from_sheet_index();
            }
            if (_AGEN_mainform.Data_Table_profile_band == null || _AGEN_mainform.Data_Table_profile_band.Rows.Count == 0)
            {
                MessageBox.Show("no sheet index defined");
                set_enable_true();
                return;
            }

            string fisier_prof = "";
            string ProjFolder = _AGEN_mainform.tpage_setup.Get_project_database_folder();

            if (System.IO.Directory.Exists(ProjFolder) == true)
            {
                if (ProjFolder.Substring(ProjFolder.Length - 1, 1) != "\\")
                {
                    ProjFolder = ProjFolder + "\\";
                }
                fisier_prof = ProjFolder + _AGEN_mainform.prof_excel_name;
                if (System.IO.File.Exists(fisier_prof) == false)
                {
                    MessageBox.Show("the profile data file does not exist");
                    set_enable_true();
                    return;
                }
                string fisier_cl = ProjFolder + _AGEN_mainform.cl_excel_name;
            }
            else
            {
                MessageBox.Show("the project database folder does not exist");
                set_enable_true();
                return;
            }


            string fisier_prof_band = ProjFolder + _AGEN_mainform.band_prof_excel_name;

            if (System.IO.File.Exists(fisier_prof_band) == true)
            {
                Functions.create_backup(fisier_prof_band);
            }


            _AGEN_mainform.Poly2D = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
            _AGEN_mainform.Poly3D = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);

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


                            Point3d pt_on_2d = _AGEN_mainform.Poly2D.GetClosestPointTo(new Point3d(x, y, 0), Vector3d.ZAxis, false);
                            double param1 = _AGEN_mainform.Poly2D.GetParameterAtPoint(pt_on_2d);
                            double eq_meas = _AGEN_mainform.Poly3D.GetDistanceAtParameter(param1);
                            _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;

                        }
                    }

                }

            }


            int elev_round = 0;
            if (Functions.IsNumeric(textBox_elev_round.Text) == true)
            {
                elev_round = Math.Abs(Convert.ToInt32(textBox_elev_round.Text));
            }

            bool rot_sta = false;
            if (checkBox_sta_at_90.Checked == true) rot_sta = true;

            double Ymin = 200000000000;

            double Xmin = -1.234;

            for (int i = 0; i < _AGEN_mainform.Data_Table_profile_band.Rows.Count; ++i)
            {
                if (_AGEN_mainform.Data_Table_profile_band.Rows[i]["y0"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[i]["y0"])) == true)
                {
                    double y1 = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[i]["y0"]);
                    if (y1 < Ymin)
                    {
                        Ymin = y1;

                        if (_AGEN_mainform.Data_Table_profile_band.Rows[i]["x0"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_profile_band.Rows[i]["x0"])) == true)
                        {
                            Xmin = Convert.ToDouble(_AGEN_mainform.Data_Table_profile_band.Rows[i]["x0"]);
                        }
                    }
                }
            }





            if (Texth == 0)
            {
                MessageBox.Show("the text style you specified does not have a set height\r\nOperation aborted");


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

                _AGEN_mainform.tpage_owner_scan.Hide();
                _AGEN_mainform.tpage_owner_draw.Hide();
                _AGEN_mainform.tpage_mat.Hide();
                _AGEN_mainform.tpage_cust_scan.Hide();
                _AGEN_mainform.tpage_cust_draw.Hide();
                _AGEN_mainform.tpage_sheet_gen.Hide();


                _AGEN_mainform.tpage_profdraw.Show();


                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;

            }


            Ag.WindowState = FormWindowState.Minimized;


            try
            {

                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = (BlockTable)ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead);
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = (BlockTableRecord)Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);


                        Autodesk.AutoCAD.EditorInput.PromptPointResult Point_res1;
                        Autodesk.AutoCAD.EditorInput.PromptPointOptions PP1;
                        PP1 = new Autodesk.AutoCAD.EditorInput.PromptPointOptions("\nPlease specify the insertion point");
                        PP1.AllowNone = false;


                        Point3d pt_ins = new Point3d(Xmin, Ymin, 0);
                        if (Ymin == 200000000000)
                        {
                            Point_res1 = Editor1.GetPoint(PP1);

                            if (Point_res1.Status != Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                            {
                                set_enable_true();
                                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                return;
                            }
                            pt_ins = Point_res1.Value;
                            Xmin = pt_ins.X;
                        }


                        double Hexag = 1;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
                        {
                            Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
                        }

                        double Vexag = 1;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex()) == true)
                        {
                            Vexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex());
                        }

                        double Hincr = 100;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hspacing()) == true)
                        {
                            Hincr = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hspacing());
                        }

                        double vincr = 100;
                        if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vspacing()) == true)
                        {
                            vincr = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vspacing());
                        }


                        System.Data.DataTable dt_top = new System.Data.DataTable();
                        System.Data.DataTable dt_prof_data = Load_existing_profile_graph(fisier_prof, ref dt_top);


                        bool left_label = true;
                        if (_AGEN_mainform.tpage_profdraw.get_comboBox_prof_el_lbl_loc() == 2)
                        {
                            left_label = false;
                        }

                        bool right_label = true;
                        if (_AGEN_mainform.tpage_profdraw.get_comboBox_prof_el_lbl_loc() == 1)
                        {
                            right_label = false;
                        }

                        string Suff = "'";
                        if (_AGEN_mainform.units_of_measurement == "m")
                        {
                            Suff = "";
                        }

                        bool draw_from_start = false;
                        if (checkBox_draw_ver_at_start.Checked == true) draw_from_start = true;
                        bool hydro = false;
                        if (checkBox_hydro_style.Checked == true)
                        {
                            hydro = true;
                        }

                        bool draw_pipe = false;
                        bool use_prof_height = false;
                        double hmax = 0;
                        if (checkBox_prof_grid_height.Checked == true && Functions.IsNumeric(textBox_prof_grid_height.Text) == true)
                        {
                            hmax = Math.Abs(Convert.ToDouble(textBox_prof_grid_height.Text));
                            use_prof_height = true;
                        }

                        bool display_match = false;
                        if (checkBox_add_matchline_label.Checked == true)
                        {
                            display_match = true;
                        }

                        textBox_overwrite_text_height.Text = Convert.ToString(Texth);
                        checkBox_overwrite_text_height.Checked = true;


                        Functions.Draw_band_profile(dt_prof_data, dt_top, pt_ins, Hincr, vincr, Hexag, Vexag,
                                                                                   _AGEN_mainform.layer_prof_grid,
                                                                                   _AGEN_mainform.layer_prof_text,
                                                                                   _AGEN_mainform.layer_prof_ground,
                                                                                   _AGEN_mainform.layer_prof_pipe,
                                                                                   _AGEN_mainform.layer_prof_smys,
                                                                                   Texth, elev_round, rot_sta,
                                                                                   Functions.Get_textstyle_id(_AGEN_mainform.tpage_profdraw.get_comboBox_prof_textstyle()),
                                                                                   Suff, left_label, right_label, _AGEN_mainform.units_of_measurement, _AGEN_mainform.Data_Table_profile_band, draw_from_start,
                                                                                   Xmin, Ymin, hydro, _AGEN_mainform.dt_station_equation, draw_pipe, checkBox_smys.Checked,
                                                                                   use_prof_height, hmax, display_match, _AGEN_mainform.config_path);




                        if (_AGEN_mainform.Poly3D != null)
                        {
                            try
                            {
                                Polyline3d poly3 = Trans1.GetObject(_AGEN_mainform.Poly3D.ObjectId, OpenMode.ForWrite) as Polyline3d;
                                poly3.Erase();
                            }
                            catch (System.Exception ex)
                            {

                            }
                        }

                        Trans1.Commit();

                        if (_AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
                        {


                            Populate_profile_band_file_with_data(fisier_prof_band, _AGEN_mainform.config_path);

                            _AGEN_mainform.Data_Table_profile_band = null;
                        }

                        _AGEN_mainform.lista_gen_prof_band = null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }




            Ag.WindowState = FormWindowState.Normal;






            set_enable_true();
        }


        private void Populate_profile_band_file_with_data(string File1, string cfg2)
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
                Microsoft.Office.Interop.Excel.Workbook Workbook2 = null;

                if (System.IO.File.Exists(File1) == false)
                {
                    Workbook1 = Excel1.Workbooks.Add();
                }

                else
                {
                    Workbook1 = Excel1.Workbooks.Open(File1);
                }
                Workbook2 = Excel1.Workbooks.Open(cfg2);
                Microsoft.Office.Interop.Excel.Worksheet W1 = Workbook1.Worksheets[1];
                Microsoft.Office.Interop.Excel.Worksheet W2 = null;

                try
                {
                    string segment1 = _AGEN_mainform.current_segment;
                    if (segment1 == "not defined") segment1 = "";

                    Functions.Transfer_to_worksheet_Data_table(W1, _AGEN_mainform.Data_Table_profile_band, _AGEN_mainform.Start_row_profile_band, "General");
                    Functions.Create_header_profile_band_file(W1, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                    if (System.IO.File.Exists(File1) == false)
                    {
                        Workbook1.SaveAs(File1);
                    }

                    else
                    {
                        Workbook1.Save();
                    }

                    Workbook1.Close();




                    foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook2.Worksheets)
                    {
                        if (wsh1.Name == "pdc2_" + segment1)
                        {
                            W2 = wsh1;
                        }
                    }

                    if (W2 == null)
                    {
                        W2 = Workbook2.Worksheets.Add(System.Reflection.Missing.Value, Workbook2.Worksheets[Workbook2.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                        W2.Name = "pdc2_" + segment1;
                    }

                    int NrR = 26;
                    int NrC = 2;


                    Object[,] values = new object[NrR, NrC];
                    values[0, 0] = "Label Text Height";
                    values[1, 0] = "X profile start";
                    values[2, 0] = "Y profile start";

                    values[3, 0] = "X elevation left";
                    values[4, 0] = "X elevation right";
                    values[5, 0] = "Y station down";
                    values[6, 0] = "Horizontal exaggeration";
                    values[6, 1] = get_textBox_prof_Hex();
                    values[7, 0] = "Vertical exaggeration";
                    values[7, 1] = get_textBox_prof_Vex();
                    values[8, 0] = "Start elevation";
                    values[8, 1] = get_textBox_prof_Elev_bottom();
                    values[9, 0] = "End elevation";
                    values[9, 1] = get_textBox_prof_Elev_top();
                    values[10, 0] = "Start station";
                    values[11, 0] = "End station";
                    values[12, 0] = "Width of the side viewports";

                    values[13, 0] = "text style:";
                    values[13, 1] = get_comboBox_prof_textstyle();


                    values[14, 0] = "horizontal spacing:";
                    values[14, 1] = get_textBox_prof_Hspacing();


                    values[15, 0] = "vertical spacing:";
                    values[15, 1] = get_textBox_prof_Vspacing();

                    values[16, 0] = "Elevation label location:";

                    if (get_comboBox_prof_el_lbl_loc() == 0)
                    {
                        values[16, 1] = "Both";
                    }
                    else if (get_comboBox_prof_el_lbl_loc() == 1)
                    {
                        values[16, 1] = "Left";
                    }
                    else if (get_comboBox_prof_el_lbl_loc() == 2)
                    {
                        values[16, 1] = "Right";
                    }




                    values[17, 0] = "elevation Rounding:";
                    values[17, 1] = _AGEN_mainform.tpage_profdraw.get_textBox_elev_round().ToString();
                    values[18, 0] = "Bottom station rotation";
                    values[18, 1] = "0";
                    values[19, 0] = "XX";
                    values[19, 1] = "XX";

                    values[20, 0] = "Draw first vertical line at start of profile line";
                    values[20, 1] = _AGEN_mainform.tpage_profdraw.get_checkBox_draw_ver_at_start().ToString();

                    values[21, 0] = "Zero = (M1+M2)/2";
                    values[21, 1] = _AGEN_mainform.tpage_profdraw.get_checkBox_set_zero_at_middle_of_profile().ToString();

                    values[22, 0] = "Hydrostatic style";
                    values[22, 1] = _AGEN_mainform.tpage_profdraw.get_checkBox_hydro_style().ToString();

                    values[23, 0] = "Display Bottom Stations at 90 Degrees";
                    values[23, 1] = _AGEN_mainform.tpage_profdraw.get_checkBox_sta_at_90().ToString();

                    values[24, 0] = "Elevation Rounding (No of decimals)";
                    values[24, 1] = get_textBox_elev_round();


                    Microsoft.Office.Interop.Excel.Range range1 = W2.Range["A1:B26"];
                    range1.Cells.NumberFormat = "General";
                    range1.Value2 = values;
                    Functions.Color_border_range_inside(range1, 0);

                    Workbook2.Save();
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
                    if (W2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W2);
                    if (Workbook1 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Workbook2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook1);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);

            }

        }

        private void button_dwg_Click(object sender, EventArgs e)
        {
            generate_profile_band_file_from_sheet_index();

            if (_AGEN_mainform.Data_Table_profile_band != null && _AGEN_mainform.Data_Table_profile_band.Rows.Count > 0)
            {
                foreach (System.Windows.Forms.Form Forma1 in System.Windows.Forms.Application.OpenForms)
                {
                    if (Forma1 is Alignment_mdi.AGEN_dwg_selection)
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
                    Alignment_mdi.AGEN_dwg_selection forma2 = new Alignment_mdi.AGEN_dwg_selection();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(forma2);
                    forma2.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - forma2.Width) / 2,
                         (Screen.PrimaryScreen.WorkingArea.Height - forma2.Height) / 2);

                }
                catch (System.Exception EX)
                {
                    MessageBox.Show(EX.Message);
                }
            }
        }



        public void set_textBox_elev_round(string txt)
        {
            textBox_elev_round.Text = txt;
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



        private void checkBox_cover_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_cover.Checked == true) checkBox_elevation.Checked = false;
        }

        private void checkBox_elevation_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_elevation.Checked == true) checkBox_cover.Checked = false;
        }

        private void button_calc_top_Click(object sender, EventArgs e)
        {
            double Vexag = 1;
            if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex()) == true)
            {
                Vexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex());
            }
            double Hexag = 1;
            if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
            {
                Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
            }

            if (Vexag != 1)
            {
                if (MessageBox.Show("the vertical exaggeration is not 1\r\ndo you want to continue?", "AGEN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            }

            if (Hexag != 1)
            {
                if (MessageBox.Show("the horizontal exaggeration is not 1\r\ndo you want to continue?", "AGEN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            }

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }



            string crossingxl = ProjF + _AGEN_mainform.crossing_excel_name;
            string profxl = ProjF + _AGEN_mainform.prof_excel_name;



            if (System.IO.File.Exists(crossingxl) == true && System.IO.File.Exists(profxl) == true)
            {
                System.Data.DataTable dt_null = null;
                System.Data.DataTable dt_prof = Load_existing_profile_graph(profxl, ref dt_null);
                if (dt_prof != null && dt_prof.Rows.Count > 0)
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
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_ground;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_ground;
                                Prompt_ground = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the ground polyline:");
                                Prompt_ground.SetRejectMessage("\nSelect a polyline!");
                                Prompt_ground.AllowNone = true;
                                Prompt_ground.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                                this.MdiParent.WindowState = FormWindowState.Minimized;

                                Rezultat_ground = ThisDrawing.Editor.GetEntity(Prompt_ground);

                                if (Rezultat_ground.Status != PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Autodesk.AutoCAD.EditorInput.PromptEntityResult Rezultat_top;
                                Autodesk.AutoCAD.EditorInput.PromptEntityOptions Prompt_top;
                                Prompt_top = new Autodesk.AutoCAD.EditorInput.PromptEntityOptions("\nSelect the top of pipe polyline:");
                                Prompt_top.SetRejectMessage("\nSelect a polyline!");
                                Prompt_top.AllowNone = true;
                                Prompt_top.AddAllowedClass(typeof(Autodesk.AutoCAD.DatabaseServices.Polyline), false);

                                this.MdiParent.WindowState = FormWindowState.Minimized;

                                Rezultat_top = ThisDrawing.Editor.GetEntity(Prompt_top);

                                if (Rezultat_top.Status != PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }



                                Polyline poly_ground = Trans1.GetObject(Rezultat_ground.ObjectId, OpenMode.ForRead) as Polyline;
                                Polyline poly_top = Trans1.GetObject(Rezultat_top.ObjectId, OpenMode.ForRead) as Polyline;

                                if (poly_ground != null && poly_top != null && poly_top.NumberOfVertices > 2)
                                {
                                    double Xstart = poly_ground.StartPoint.X;
                                    double Xend = poly_ground.EndPoint.X;
                                    double Ystart = poly_ground.StartPoint.Y;

                                    if (dt_prof.Rows[0][_AGEN_mainform.Col_station] != DBNull.Value && dt_prof.Rows[dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station] != DBNull.Value && dt_prof.Rows[0][_AGEN_mainform.Col_Elev] != DBNull.Value)
                                    {
                                        double Sta_start = Convert.ToDouble(dt_prof.Rows[0][_AGEN_mainform.Col_station]);
                                        double Elev_start = Convert.ToDouble(dt_prof.Rows[0][_AGEN_mainform.Col_Elev]);
                                        double Sta_end = Convert.ToDouble(dt_prof.Rows[dt_prof.Rows.Count - 1][_AGEN_mainform.Col_station]);

                                        if (Math.Round(Math.Abs(Xstart - Xend) * Hexag, 2) == Math.Round(Sta_end - Sta_start, 2))
                                        {
                                            double max_bend = 1000;
                                            if (Functions.IsNumeric(textBox_max_bend.Text) == true)
                                            {
                                                max_bend = Convert.ToDouble(textBox_max_bend.Text);
                                            }

                                            int round1 = 0;
                                            if (Functions.IsNumeric(textBox_ang_round.Text) == true)
                                            {
                                                round1 = Convert.ToInt32(textBox_ang_round.Text);
                                            }

                                            System.Data.DataTable dt_xing = Functions.Creaza_crossing_datatable_structure();
                                            System.Data.DataTable dt1 = dt_prof.Clone();

                                            for (int i = 1; i < poly_top.NumberOfVertices - 1; ++i)
                                            {
                                                Point3d pt1 = poly_top.GetPointAtParameter(i - 1);
                                                Point3d pt2 = poly_top.GetPointAtParameter(i);
                                                Point3d pt3 = poly_top.GetPointAtParameter(i + 1);

                                                if (i == 1)
                                                {
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_station] = Sta_start + Math.Abs(pt1.X - Xstart) / Hexag;
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Elev] = Elev_start - (poly_ground.StartPoint.Y - pt1.Y) / Vexag;
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";
                                                }


                                                double sta = Sta_start + Math.Abs(pt2.X - Xstart) / Hexag;
                                                double elev = Elev_start - (poly_ground.StartPoint.Y - pt2.Y) / Vexag;

                                                dt1.Rows.Add();
                                                dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_station] = sta;
                                                dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Elev] = elev;
                                                dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";

                                                if (i == poly_top.NumberOfVertices - 2)
                                                {
                                                    dt1.Rows.Add();
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_station] = Sta_start + Math.Abs(pt3.X - Xstart) / Hexag;
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Elev] = Elev_start - (poly_ground.StartPoint.Y - pt3.Y) / Vexag;
                                                    dt1.Rows[dt1.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";
                                                }

                                                double defl_rad = Functions.Get_deflection_angle_rad(Xstart + (pt1.X - Xstart) / Hexag, Ystart - (Ystart - pt1.Y) / Vexag,
                                                                                                     Xstart + (pt2.X - Xstart) / Hexag, Ystart - (Ystart - pt2.Y) / Vexag,
                                                                                                     Xstart + (pt3.X - Xstart) / Hexag, Ystart - (Ystart - pt3.Y) / Vexag);
                                                string defl_side = Functions.Get_deflection_side(pt1.X, pt1.Y, pt2.X, pt2.Y, pt3.X, pt3.Y).Replace("LT", "SB").Replace("RT", "OB");
                                                double d_dms = 180 * defl_rad / Math.PI;

                                                string prefix = "P.I. < ";
                                                if (d_dms > max_bend)
                                                {
                                                    prefix = "I.B. < ";
                                                }

                                                dt_xing.Rows.Add();
                                                dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.col_desc] = prefix + Convert.ToString(Math.Round(d_dms, round1)) + "° " + defl_side;
                                                dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_Elev] = elev;
                                                dt_xing.Rows[dt_xing.Rows.Count - 1]["DispProf"] = "YES";
                                                if (_AGEN_mainform.Project_type == "2D")
                                                {
                                                    dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = sta;
                                                }
                                                else
                                                {
                                                    dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = sta;
                                                }
                                            }


                                            Microsoft.Office.Interop.Excel.Worksheet W_xing = null;
                                            Microsoft.Office.Interop.Excel.Worksheet W_prof = null;
                                            Microsoft.Office.Interop.Excel.Application Excel1 = null;
                                            Microsoft.Office.Interop.Excel.Workbook Workboook_xing = null;
                                            Microsoft.Office.Interop.Excel.Workbook Workboook_prof = null;

                                            bool is_opened_xing = false;
                                            bool save_and_close_xing = false;
                                            bool is_opened_prof = false;
                                            bool save_and_close_prof = false;
                                            try
                                            {
                                                try
                                                {
                                                    Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                                                    foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                                                    {

                                                        if (Workbook2.FullName == crossingxl)
                                                        {
                                                            Workboook_xing = Workbook2;
                                                            is_opened_xing = true;
                                                        }
                                                        if (Workbook2.FullName == profxl)
                                                        {
                                                            Workboook_prof = Workbook2;
                                                            is_opened_prof = true;
                                                        }

                                                    }
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    Excel1 = new Microsoft.Office.Interop.Excel.Application();
                                                }


                                                if (is_opened_xing == false)
                                                {

                                                    Workboook_xing = Excel1.Workbooks.Open(crossingxl);
                                                    save_and_close_xing = true;

                                                }
                                                if (is_opened_prof == false)
                                                {
                                                    Workboook_prof = Excel1.Workbooks.Open(profxl);
                                                    save_and_close_prof = true;
                                                }

                                                List<string> lista_xing = new List<string>();
                                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workboook_xing.Worksheets)
                                                {
                                                    if (Wx.Name.Contains("prof_lab") == true)
                                                    {
                                                        lista_xing.Add(Wx.Name);
                                                    }
                                                }

                                                List<string> lista_prof = new List<string>();
                                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workboook_prof.Worksheets)
                                                {
                                                    lista_prof.Add(Wx.Name);
                                                }

                                                string new_name_xing_orig = "prof_lab";
                                                string new_xing_name = "prof_lab";
                                                int index1 = 0;
                                                if (lista_xing.Contains(new_xing_name) == true)
                                                {
                                                    do
                                                    {
                                                        ++index1;
                                                        new_xing_name = new_name_xing_orig + Convert.ToString(index1);
                                                    } while (lista_xing.Contains(new_xing_name) == true);
                                                }
                                                W_xing = Workboook_xing.Worksheets.Add(System.Reflection.Missing.Value, Workboook_xing.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                                W_xing.Name = new_xing_name;

                                                string new_name_prof_orig = "TOP";
                                                string new_prof_name = "TOP";
                                                int index2 = 0;
                                                if (lista_prof.Contains(new_prof_name) == true)
                                                {
                                                    do
                                                    {
                                                        ++index2;
                                                        new_prof_name = new_name_prof_orig + Convert.ToString(index1);
                                                    } while (lista_prof.Contains(new_prof_name) == true);
                                                }

                                                foreach (Microsoft.Office.Interop.Excel.Worksheet Wx in Workboook_prof.Worksheets)
                                                {
                                                    if (Wx.Name == new_name_prof_orig)
                                                    {
                                                        Wx.Name = new_prof_name;
                                                    }
                                                }

                                                W_prof = Workboook_prof.Worksheets.Add(System.Reflection.Missing.Value, Workboook_prof.Worksheets[1], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                                                W_prof.Name = new_name_prof_orig;

                                                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                                if (segment1 == "not defined") segment1 = "";

                                                if (dt_xing.Rows.Count > 0)
                                                {

                                                    Functions.Create_header_crossing_file(W_xing, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                                                    W_xing.Cells.NumberFormat = "General";
                                                    int maxRows = dt_xing.Rows.Count;
                                                    int maxCols = dt_xing.Columns.Count;


                                                    Microsoft.Office.Interop.Excel.Range range1 = W_xing.Range["A9:R" + (9 + maxRows - 1).ToString()];




                                                    object[,] values1 = new object[maxRows, maxCols];

                                                    for (int i = 0; i < maxRows; ++i)
                                                    {
                                                        for (int j = 0; j < maxCols; ++j)
                                                        {
                                                            if (dt_xing.Rows[i][j] != DBNull.Value)
                                                            {
                                                                values1[i, j] = dt_xing.Rows[i][j];
                                                            }
                                                        }
                                                    }
                                                    range1.Value2 = values1;
                                                }


                                                if (dt1.Rows.Count > 0)
                                                {
                                                    Functions.Create_header_graph_profile(W_prof, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                                                    W_prof.Cells.NumberFormat = "General";
                                                    int maxRows = dt1.Rows.Count;
                                                    int maxCols = dt1.Columns.Count;


                                                    Microsoft.Office.Interop.Excel.Range range0 = W_prof.Range["A8:G8"];
                                                    object[,] values0 = new object[1, maxCols];


                                                    for (int j = 0; j < maxCols; ++j)
                                                    {
                                                        values0[0, j] = dt1.Columns[j].ColumnName;
                                                    }
                                                    range0.Value2 = values0;



                                                    Microsoft.Office.Interop.Excel.Range range1 = W_prof.Range["A9:G" + (9 + maxRows - 1).ToString()];
                                                    object[,] values1 = new object[maxRows, maxCols];

                                                    for (int i = 0; i < maxRows; ++i)
                                                    {
                                                        for (int j = 0; j < maxCols; ++j)
                                                        {
                                                            if (dt1.Rows[i][j] != DBNull.Value)
                                                            {
                                                                values1[i, j] = dt1.Rows[i][j];
                                                            }
                                                        }
                                                    }
                                                    range1.Value2 = values1;
                                                }

                                                Workboook_xing.Save();
                                                Workboook_prof.Save();

                                                if (save_and_close_xing == true)
                                                {
                                                    Workboook_xing.Close();
                                                }


                                                if (save_and_close_prof == true)
                                                {
                                                    Workboook_prof.Close();
                                                }

                                            }
                                            catch (System.Exception ex)
                                            {
                                                System.Windows.Forms.MessageBox.Show(ex.Message);
                                            }
                                            finally
                                            {
                                                if (W_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_xing);
                                                if (Workboook_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workboook_xing);
                                                if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
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

                    this.MdiParent.WindowState = FormWindowState.Normal;
                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                    set_enable_true();

                }
            }

        }

        private void button_top_from_profile_band_Click(object sender, EventArgs e)
        {

            string xing_tab = comboBox_xing_tab.Text;
            string prof_tab = comboBox_profile_tab.Text;

            if (xing_tab == "" || prof_tab == "")
            {
                MessageBox.Show("no crossing/profile tab specifed", "AGEN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string Agen_profile_band_V2 = "Agen_profile_band_V2";
            string Agen_profile_band_V3 = "Agen_profile_band_V3";
            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;

            Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (Tables1.IsTableDefined(Agen_profile_band_V2) == false && Tables1.IsTableDefined(Agen_profile_band_V3) == false)
            {
                MessageBox.Show("no " + Agen_profile_band_V3 + " data table defined\r\noperation aborted", "AGEN", MessageBoxButtons.OK, MessageBoxIcon.Error);

                set_enable_true();
                ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                return;
            }


            double Vexag = 1;
            if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex()) == true)
            {
                Vexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Vex());
            }
            double Hexag = 1;
            if (Functions.IsNumeric(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex()) == true)
            {
                Hexag = Convert.ToDouble(_AGEN_mainform.tpage_profdraw.get_textBox_prof_Hex());
            }

            if (Vexag != 1)
            {
                if (MessageBox.Show("the vertical exaggeration is not 1\r\ndo you want to continue?", "AGEN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            }

            if (Hexag != 1)
            {
                if (MessageBox.Show("the horizontal exaggeration is not 1\r\ndo you want to continue?", "AGEN", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
            }

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }



            string crossingxl = ProjF + _AGEN_mainform.crossing_excel_name;
            string profxl = ProjF + _AGEN_mainform.prof_excel_name;



            if (System.IO.File.Exists(crossingxl) == true && System.IO.File.Exists(profxl) == true)
            {
                System.Data.DataTable dt_null = null;
                System.Data.DataTable dt_prof = Load_existing_profile_graph(profxl, ref dt_null);
                if (dt_prof != null && dt_prof.Rows.Count > 0)
                {

                    ObjectId[] Empty_array = null;

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
                                BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;





                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_ground;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_ground = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_ground.MessageForAdding = "\nSelect the ground polylines:";
                                Prompt_ground.SingleOnly = false;


                                this.MdiParent.WindowState = FormWindowState.Minimized;

                                Rezultat_ground = ThisDrawing.Editor.GetSelection(Prompt_ground);

                                if (Rezultat_ground.Status != PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                Autodesk.AutoCAD.EditorInput.PromptSelectionResult Rezultat_top;
                                Autodesk.AutoCAD.EditorInput.PromptSelectionOptions Prompt_top = new Autodesk.AutoCAD.EditorInput.PromptSelectionOptions();
                                Prompt_top.MessageForAdding = "\nSelect the top of pipe polyline:";
                                Prompt_top.SingleOnly = false;

                                Rezultat_top = ThisDrawing.Editor.GetSelection(Prompt_top);

                                if (Rezultat_top.Status != PromptStatus.OK)
                                {
                                    this.MdiParent.WindowState = FormWindowState.Normal;
                                    set_enable_true();
                                    ThisDrawing.Editor.WriteMessage("\n" + "Command:");
                                    return;
                                }

                                this.MdiParent.WindowState = FormWindowState.Normal;

                                List<Polyline> lista_ground = new List<Polyline>();
                                List<Polyline> lista_top = new List<Polyline>();


                                for (int i = 0; i < Rezultat_ground.Value.Count; ++i)
                                {
                                    Polyline poly1 = Trans1.GetObject(Rezultat_ground.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                    if (poly1 != null)
                                    {
                                        lista_ground.Add(poly1);
                                    }

                                }
                                for (int i = 0; i < Rezultat_top.Value.Count; ++i)
                                {
                                    Polyline poly1 = Trans1.GetObject(Rezultat_top.Value[i].ObjectId, OpenMode.ForRead) as Polyline;
                                    if (poly1 != null)
                                    {
                                        lista_top.Add(poly1);
                                    }

                                    Line line1 = Trans1.GetObject(Rezultat_top.Value[i].ObjectId, OpenMode.ForRead) as Line;
                                    if (line1 != null)
                                    {
                                        Polyline pl1 = new Polyline();
                                        pl1.AddVertexAt(0, new Point2d(line1.StartPoint.X, line1.StartPoint.Y), 0, 0, 0);
                                        pl1.AddVertexAt(1, new Point2d(line1.EndPoint.X, line1.EndPoint.Y), 0, 0, 0);
                                        lista_top.Add(pl1);
                                    }
                                }

                                string segm2 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                if (_AGEN_mainform.tpage_setup.Get_segment_name1() == "not defined")
                                {
                                    segm2 = "";
                                }


                                for (int j = lista_ground.Count - 1; j >= 0; --j)
                                {
                                    Polyline poly2 = lista_ground[j];

                                    double start1 = -123.4;
                                    double end1 = -123.4;
                                    string segm1 = "123456";

                                    if (Tables1.GetTableNames().Contains(Agen_profile_band_V2) == true)
                                    {
                                        using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V2])
                                        {

                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), poly2.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                            {
                                                if (Records1.Count > 0)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {

                                                        for (int k = 0; k < Record1.Count; ++k)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = Record1[k].StrValue;

                                                            if (Nume_field.ToLower() == "beginsta")
                                                            {
                                                                if (Functions.IsNumeric(Valoare_field) == true)
                                                                {
                                                                    start1 = Convert.ToDouble(Valoare_field);
                                                                }
                                                            }

                                                            if (Nume_field.ToLower() == "endsta")
                                                            {
                                                                if (Functions.IsNumeric(Valoare_field) == true)
                                                                {
                                                                    end1 = Convert.ToDouble(Valoare_field);
                                                                }
                                                            }
                                                            if (Nume_field.ToLower() == "segment")
                                                            {
                                                                segm1 = Convert.ToString(Valoare_field);
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }

                                    if (Tables1.GetTableNames().Contains(Agen_profile_band_V3) == true)
                                    {
                                        using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V3])
                                        {

                                            using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), poly2.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                            {
                                                if (Records1.Count > 0)
                                                {
                                                    Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                    foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                    {

                                                        for (int k = 0; k < Record1.Count; ++k)
                                                        {
                                                            Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                            string Nume_field = Field_def1.Name;
                                                            string Valoare_field = Record1[k].StrValue;

                                                            if (Nume_field.ToLower() == "beginsta")
                                                            {
                                                                if (Functions.IsNumeric(Valoare_field) == true)
                                                                {
                                                                    start1 = Convert.ToDouble(Valoare_field);
                                                                }
                                                            }

                                                            if (Nume_field.ToLower() == "endsta")
                                                            {
                                                                if (Functions.IsNumeric(Valoare_field) == true)
                                                                {
                                                                    end1 = Convert.ToDouble(Valoare_field);
                                                                }
                                                            }
                                                            if (Nume_field.ToLower() == "segment")
                                                            {
                                                                segm1 = Convert.ToString(Valoare_field);
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    if (segm1.ToLower() != segm2.ToLower() || start1 == -123.4 || end1 == -123.4)
                                    {
                                        lista_ground.RemoveAt(j);
                                    }

                                }

                                if (lista_ground.Count == 0)
                                {
                                    MessageBox.Show("no ground polyline selected", "AGEN", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }

                                System.Data.DataTable dt_xing = Functions.Creaza_crossing_datatable_structure();
                                System.Data.DataTable dt_top = dt_prof.Clone();


                                for (int s = 0; s < lista_top.Count; ++s)
                                {

                                    Polyline poly_top = lista_top[s];
                                    Point3d p1 = poly_top.StartPoint;

                                    double dist1 = 200000;
                                    Polyline poly_ground = null;

                                    for (int j = 0; j < lista_ground.Count; ++j)
                                    {
                                        Polyline poly2 = lista_ground[j];
                                        Point3d p2 = poly2.GetClosestPointTo(new Point3d(p1.X, p1.Y, poly2.Elevation), Vector3d.ZAxis, false);

                                        double dist2 = Math.Pow(Math.Pow((p1.X - p2.X), 2) + Math.Pow((p1.Y - p2.Y), 2), 0.5);

                                        if (dist2 < dist1)
                                        {
                                            poly_ground = poly2;
                                            dist1 = dist2;
                                        }

                                    }

                                    if (poly_ground != null && poly_top != null && poly_top.NumberOfVertices > 2)
                                    {

                                        double Sta_start = -123.4;
                                        double Sta_end = -123.4;
                                        if (Tables1.GetTableNames().Contains(Agen_profile_band_V2) == true)
                                        {
                                            using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V2])
                                            {

                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), poly_ground.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                {
                                                    if (Records1.Count > 0)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {

                                                            for (int k = 0; k < Record1.Count; ++k)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                                string Nume_field = Field_def1.Name;
                                                                string Valoare_field = Record1[k].StrValue;

                                                                if (Nume_field.ToLower() == "beginsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        Sta_start = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }

                                                                if (Nume_field.ToLower() == "endsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        Sta_end = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                        if (Tables1.GetTableNames().Contains(Agen_profile_band_V3) == true)
                                        {
                                            using (Autodesk.Gis.Map.ObjectData.Table Tabla1 = Tables1[Agen_profile_band_V3])
                                            {

                                                using (Autodesk.Gis.Map.ObjectData.Records Records1 = Tabla1.GetObjectTableRecords(Convert.ToUInt32(0), poly_ground.ObjectId, Autodesk.Gis.Map.Constants.OpenMode.OpenForRead, true))
                                                {
                                                    if (Records1.Count > 0)
                                                    {
                                                        Autodesk.Gis.Map.ObjectData.FieldDefinitions Field_defs1 = Tabla1.FieldDefinitions;
                                                        foreach (Autodesk.Gis.Map.ObjectData.Record Record1 in Records1)
                                                        {

                                                            for (int k = 0; k < Record1.Count; ++k)
                                                            {
                                                                Autodesk.Gis.Map.ObjectData.FieldDefinition Field_def1 = Field_defs1[k];
                                                                string Nume_field = Field_def1.Name;
                                                                string Valoare_field = Record1[k].StrValue;

                                                                if (Nume_field.ToLower() == "beginsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        Sta_start = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }

                                                                if (Nume_field.ToLower() == "endsta")
                                                                {
                                                                    if (Functions.IsNumeric(Valoare_field) == true)
                                                                    {
                                                                        Sta_end = Convert.ToDouble(Valoare_field);
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                        double Xstart = poly_ground.StartPoint.X;
                                        double Xend = poly_ground.EndPoint.X;
                                        double Ystart = poly_ground.StartPoint.Y;
                                        double Elev_start = -1.234;


                                        for (int j = 1; j < dt_prof.Rows.Count; ++j)
                                        {
                                            if (dt_prof.Rows[j][_AGEN_mainform.Col_station] != DBNull.Value && dt_prof.Rows[j][_AGEN_mainform.Col_Elev] != DBNull.Value)
                                            {
                                                double elev1 = Convert.ToDouble(dt_prof.Rows[j - 1][_AGEN_mainform.Col_Elev]);
                                                double elev2 = Convert.ToDouble(dt_prof.Rows[j][_AGEN_mainform.Col_Elev]);

                                                double sta1 = Convert.ToDouble(dt_prof.Rows[j - 1][_AGEN_mainform.Col_station]);
                                                double sta2 = Convert.ToDouble(dt_prof.Rows[j][_AGEN_mainform.Col_station]);

                                                if (sta1 <= Sta_start && Sta_start <= sta2)
                                                {
                                                    if (Math.Round(sta2, 2) == Math.Round(sta1, 2))
                                                    {
                                                        Elev_start = elev1;
                                                    }
                                                    else
                                                    {
                                                        Elev_start = elev1 + ((elev2 - elev1) * (Sta_start - sta1)) / (sta2 - sta1);
                                                    }

                                                    j = dt_prof.Rows.Count;
                                                }


                                            }
                                        }

                                        if (Elev_start != -1.234)
                                        {



                                            if (Math.Round(Math.Abs(Xstart - Xend) * Hexag, 2) == Math.Round(Sta_end - Sta_start, 2))
                                            {
                                                double max_bend = 1000;
                                                if (Functions.IsNumeric(textBox_max_bend.Text) == true)
                                                {
                                                    max_bend = Convert.ToDouble(textBox_max_bend.Text);
                                                }

                                                int round1 = 0;
                                                if (Functions.IsNumeric(textBox_ang_round.Text) == true)
                                                {
                                                    round1 = Convert.ToInt32(textBox_ang_round.Text);
                                                }



                                                for (int i = 1; i < poly_top.NumberOfVertices - 1; ++i)
                                                {
                                                    Point3d pt1 = poly_top.GetPointAtParameter(i - 1);
                                                    Point3d pt2 = poly_top.GetPointAtParameter(i);
                                                    Point3d pt3 = poly_top.GetPointAtParameter(i + 1);

                                                    if (i == 1)
                                                    {
                                                        dt_top.Rows.Add();
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_station] = Sta_start + Math.Abs(pt1.X - Xstart) / Hexag;
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Elev] = Elev_start - (poly_ground.StartPoint.Y - pt1.Y) / Vexag;
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";
                                                    }


                                                    double sta = Sta_start + Math.Abs(pt2.X - Xstart) / Hexag;
                                                    double elev = Elev_start - (poly_ground.StartPoint.Y - pt2.Y) / Vexag;

                                                    dt_top.Rows.Add();
                                                    dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_station] = sta;
                                                    dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Elev] = elev;
                                                    dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";

                                                    if (i == poly_top.NumberOfVertices - 2)
                                                    {
                                                        dt_top.Rows.Add();
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_station] = Sta_start + Math.Abs(pt3.X - Xstart) / Hexag;
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Elev] = Elev_start - (poly_ground.StartPoint.Y - pt3.Y) / Vexag;
                                                        dt_top.Rows[dt_top.Rows.Count - 1][_AGEN_mainform.Col_Type] = "TOP";
                                                    }

                                                    double defl_rad = Functions.Get_deflection_angle_rad(Xstart + (pt1.X - Xstart) / Hexag, Ystart - (Ystart - pt1.Y) / Vexag,
                                                                                                         Xstart + (pt2.X - Xstart) / Hexag, Ystart - (Ystart - pt2.Y) / Vexag,
                                                                                                         Xstart + (pt3.X - Xstart) / Hexag, Ystart - (Ystart - pt3.Y) / Vexag);
                                                    string defl_side = Functions.Get_deflection_side(pt1.X, pt1.Y, pt2.X, pt2.Y, pt3.X, pt3.Y).Replace("LT", "SB").Replace("RT", "OB");
                                                    double d_dms = 180 * defl_rad / Math.PI;

                                                    string prefix = "P.I. < ";
                                                    if (d_dms > max_bend)
                                                    {
                                                        prefix = "I.B. < ";
                                                    }

                                                    dt_xing.Rows.Add();
                                                    dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.col_desc] = prefix + Convert.ToString(Math.Round(d_dms, round1)) + "° " + defl_side;
                                                    dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_Elev] = elev;
                                                    dt_xing.Rows[dt_xing.Rows.Count - 1]["DispProf"] = "YES";
                                                    if (_AGEN_mainform.Project_type == "2D")
                                                    {
                                                        dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_2DSta] = sta;
                                                    }
                                                    else
                                                    {
                                                        dt_xing.Rows[dt_xing.Rows.Count - 1][_AGEN_mainform.Col_3DSta] = sta;
                                                    }
                                                }


                                            }
                                        }
                                    }


                                }

                                Microsoft.Office.Interop.Excel.Worksheet W_xing = null;
                                Microsoft.Office.Interop.Excel.Worksheet W_prof = null;
                                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                                Microsoft.Office.Interop.Excel.Workbook Workboook_xing = null;
                                Microsoft.Office.Interop.Excel.Workbook Workboook_prof = null;

                                bool is_opened_xing = false;
                                bool save_and_close_xing = false;
                                bool is_opened_prof = false;
                                bool save_and_close_prof = false;
                                try
                                {
                                    try
                                    {
                                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                                        {
                                            string workbookname = Workbook2.FullName;

                                            if (workbookname.ToLower() == crossingxl.ToLower())
                                            {
                                                Workboook_xing = Workbook2;
                                                is_opened_xing = true;
                                            }
                                            if (workbookname.ToLower() == profxl.ToLower())
                                            {
                                                Workboook_prof = Workbook2;
                                                is_opened_prof = true;
                                            }

                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                                    }


                                    if (is_opened_xing == false)
                                    {

                                        Workboook_xing = Excel1.Workbooks.Open(crossingxl);
                                        save_and_close_xing = true;

                                    }
                                    if (is_opened_prof == false)
                                    {
                                        Workboook_prof = Excel1.Workbooks.Open(profxl);
                                        save_and_close_prof = true;
                                    }



                                    W_xing = Workboook_xing.Worksheets[comboBox_xing_tab.Text];

                                    //  Load crossigs and add them to the dt_xing

                                    Functions.create_backup(crossingxl);
                                    System.Data.DataTable dt_ex_xing = Load_existing_crossing(W_xing);

                                    if (dt_ex_xing != null && dt_ex_xing.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dt_ex_xing.Rows.Count; ++i)
                                        {
                                            System.Data.DataRow row1 = dt_xing.NewRow();
                                            System.Data.DataRow row2 = dt_ex_xing.Rows[i];
                                            row1.ItemArray = row2.ItemArray;
                                            dt_xing.Rows.Add(row1);
                                        }

                                        dt_ex_xing = null;
                                    }


                                    W_prof = Workboook_prof.Worksheets[comboBox_profile_tab.Text];
                                    Functions.create_backup(profxl);
                                    System.Data.DataTable dt_ex_prof = Load_existing_profile(W_prof);

                                    if (dt_ex_prof != null && dt_ex_prof.Rows.Count > 0)
                                    {
                                        for (int i = 0; i < dt_ex_prof.Rows.Count; ++i)
                                        {
                                            System.Data.DataRow row1 = dt_top.NewRow();
                                            System.Data.DataRow row2 = dt_ex_prof.Rows[i];
                                            row1.ItemArray = row2.ItemArray;
                                            dt_top.Rows.Add(row1);
                                        }

                                        dt_ex_prof = null;
                                    }

                                    if (comboBox_profile_tab.Text != "TOP")
                                    {
                                        List<string> lista_sheets = new List<string>();
                                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Excel1.Worksheets)
                                        {
                                            string old_name = sheet.Name;
                                            if (old_name.ToUpper().Contains("TOP") == true)
                                            {
                                                lista_sheets.Add(old_name);
                                            }

                                        }
                                        if (lista_sheets.Count > 0)
                                        {
                                            for (int i = 0; i < lista_sheets.Count; ++i)
                                            {
                                                if (lista_sheets[i].ToUpper() == "TOP")
                                                {
                                                    int k = 1;
                                                    string new_name = "OLD_TOP" + k.ToString();
                                                    do
                                                    {
                                                        if (lista_sheets.Contains(new_name) == true)
                                                        {
                                                            ++k;
                                                            new_name = "OLD_TOP" + k.ToString();
                                                        }

                                                    } while (lista_sheets.Contains(new_name) == true);

                                                    Workboook_prof.Worksheets["TOP"].Name = new_name;
                                                    comboBox_profile_tab.Items[comboBox_profile_tab.SelectedIndex] = new_name;
                                                }
                                            }
                                        }



                                    }






                                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                    if (segment1 == "not defined") segment1 = "";

                                    if (dt_xing.Rows.Count > 0)
                                    {

                                        Microsoft.Office.Interop.Excel.Range range2 = W_xing.Range["A9:R50000"];
                                        range2.ClearContents();
                                        range2.ClearFormats();

                                        Functions.Create_header_crossing_file(W_xing, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);

                                        W_xing.Cells.NumberFormat = "General";
                                        int maxRows = dt_xing.Rows.Count;
                                        int maxCols = dt_xing.Columns.Count;


                                        Microsoft.Office.Interop.Excel.Range range1 = W_xing.Range["A9:R" + (9 + maxRows - 1).ToString()];

                                        object[,] values1 = new object[maxRows, maxCols];

                                        for (int i = 0; i < maxRows; ++i)
                                        {
                                            for (int j = 0; j < maxCols; ++j)
                                            {
                                                if (dt_xing.Rows[i][j] != DBNull.Value)
                                                {
                                                    values1[i, j] = dt_xing.Rows[i][j];
                                                }
                                            }
                                        }
                                        range1.Value2 = values1;
                                    }


                                    if (dt_top.Rows.Count > 0)
                                    {
                                        Microsoft.Office.Interop.Excel.Range range2 = W_prof.Range["A9:G50000"];
                                        range2.ClearContents();
                                        range2.ClearFormats();
                                        Functions.Create_header_graph_profile(W_prof, _AGEN_mainform.tpage_setup.Get_client_name(), _AGEN_mainform.tpage_setup.Get_project_name(), segment1);
                                        W_prof.Cells.NumberFormat = "General";
                                        W_prof.Name = "TOP";
                                        int maxRows = dt_top.Rows.Count;
                                        int maxCols = dt_top.Columns.Count;


                                        Microsoft.Office.Interop.Excel.Range range0 = W_prof.Range["A8:G8"];
                                        object[,] values0 = new object[1, maxCols];


                                        for (int j = 0; j < maxCols; ++j)
                                        {
                                            values0[0, j] = dt_top.Columns[j].ColumnName;
                                        }
                                        range0.Value2 = values0;



                                        Microsoft.Office.Interop.Excel.Range range1 = W_prof.Range["A9:G" + (9 + maxRows - 1).ToString()];
                                        object[,] values1 = new object[maxRows, maxCols];

                                        for (int i = 0; i < maxRows; ++i)
                                        {
                                            for (int j = 0; j < maxCols; ++j)
                                            {
                                                if (dt_top.Rows[i][j] != DBNull.Value)
                                                {
                                                    values1[i, j] = dt_top.Rows[i][j];
                                                }
                                            }
                                        }
                                        range1.Value2 = values1;
                                    }


                                    comboBox_profile_tab.Items.Clear();
                                    comboBox_xing_tab.Items.Clear();

                                    Workboook_xing.Save();
                                    Workboook_prof.Save();

                                    if (save_and_close_xing == true)
                                    {
                                        Workboook_xing.Close();
                                    }


                                    if (save_and_close_prof == true)
                                    {
                                        Workboook_prof.Close();
                                    }

                                }
                                catch (System.Exception ex)
                                {
                                    System.Windows.Forms.MessageBox.Show(ex.Message);
                                }
                                finally
                                {
                                    if (W_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_xing);
                                    if (Workboook_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workboook_xing);
                                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                                }




                            }
                        }
                    }

                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    this.MdiParent.WindowState = FormWindowState.Normal;
                    Editor1.SetImpliedSelection(Empty_array);
                    Editor1.WriteMessage("\nCommand:");
                    set_enable_true();

                }
            }

        }

        private void button_load_tabs_cross_and_xing_Click(object sender, EventArgs e)
        {
            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }



            string crossingxl = ProjF + _AGEN_mainform.crossing_excel_name;
            string profxl = ProjF + _AGEN_mainform.prof_excel_name;



            if (System.IO.File.Exists(crossingxl) == true && System.IO.File.Exists(profxl) == true)
            {
                Microsoft.Office.Interop.Excel.Worksheet W_xing = null;
                Microsoft.Office.Interop.Excel.Worksheet W_prof = null;
                Microsoft.Office.Interop.Excel.Application Excel1 = null;
                Microsoft.Office.Interop.Excel.Workbook Workboook_xing = null;
                Microsoft.Office.Interop.Excel.Workbook Workboook_prof = null;

                bool is_opened_xing = false;
                bool close_xing = false;
                bool is_opened_prof = false;
                bool close_prof = false;
                try
                {
                    try
                    {
                        Excel1 = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                        foreach (Microsoft.Office.Interop.Excel.Workbook Workbook2 in Excel1.Workbooks)
                        {
                            string workbookname = Workbook2.FullName;

                            if (workbookname.ToLower() == crossingxl.ToLower())
                            {
                                Workboook_xing = Workbook2;
                                is_opened_xing = true;
                            }
                            if (workbookname.ToLower() == profxl.ToLower())
                            {
                                Workboook_prof = Workbook2;
                                is_opened_prof = true;
                            }

                        }
                    }
                    catch (System.Exception ex)
                    {
                        Excel1 = new Microsoft.Office.Interop.Excel.Application();
                    }


                    if (is_opened_xing == false)
                    {

                        Workboook_xing = Excel1.Workbooks.Open(crossingxl);
                        close_xing = true;

                    }
                    if (is_opened_prof == false)
                    {
                        Workboook_prof = Excel1.Workbooks.Open(profxl);
                        close_prof = true;
                    }

                    comboBox_profile_tab.Items.Clear();
                    comboBox_xing_tab.Items.Clear();

                    for (int i = 1; i <= Workboook_prof.Worksheets.Count; ++i)
                    {
                        comboBox_profile_tab.Items.Add(Workboook_prof.Worksheets[i].Name);
                    }
                    for (int i = 1; i <= Workboook_xing.Worksheets.Count; ++i)
                    {
                        comboBox_xing_tab.Items.Add(Workboook_xing.Worksheets[i].Name);
                    }

                    if (close_prof == true)
                    {
                        Workboook_prof.Close();
                    }
                    if (close_xing == true)
                    {
                        Workboook_xing.Close();
                    }

                    if (Excel1.Workbooks.Count == 0) Excel1.Quit();

                    if (comboBox_profile_tab.Items.Count > 0)
                    {
                        comboBox_profile_tab.SelectedIndex = 0;
                    }

                    if (comboBox_xing_tab.Items.Count > 0)
                    {
                        comboBox_xing_tab.SelectedIndex = 0;
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (W_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_xing);
                    if (W_prof != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(W_prof);
                    if (Workboook_xing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workboook_xing);
                    if (Workboook_prof != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(Workboook_prof);
                    if (Excel1 != null && Excel1.Workbooks.Count == 0) System.Runtime.InteropServices.Marshal.ReleaseComObject(Excel1);
                }
            }
        }




        public System.Data.DataTable Load_existing_crossing(Microsoft.Office.Interop.Excel.Worksheet W1)
        {
            return Functions.Build_Data_table_crossings_from_excel(W1, _AGEN_mainform.Start_row_crossing + 1);
        }

        public System.Data.DataTable Load_existing_profile(Microsoft.Office.Interop.Excel.Worksheet W1)
        {

            return Functions.Build_Data_table_profile_from_excel(W1, _AGEN_mainform.Start_row_graph_profile + 1);




        }

    }
}