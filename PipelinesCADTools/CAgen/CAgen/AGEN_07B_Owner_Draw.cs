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
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;


namespace Alignment_mdi
{
    public partial class AGEN_OwnershipDraw : Form
    {

        bool Block_loaded_from_open_config = false;

        private void set_enable_false(object sender)
        {
            List<System.Windows.Forms.Button> lista_butoane = new List<Button>();
            lista_butoane.Add(button_insert_prop_band);

            lista_butoane.Add(button_show_ownership_settings);


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
            lista_butoane.Add(button_insert_prop_band);
            lista_butoane.Add(button_show_ownership_settings);
            foreach (System.Windows.Forms.Button bt1 in lista_butoane)
            {
                bt1.Enabled = true;
            }
        }


        public AGEN_OwnershipDraw()
        {
            InitializeComponent();
        }

        private void button_show_ownership_scan_Click(object sender, EventArgs e)
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

            _AGEN_mainform.tpage_owner_draw.Hide();
            _AGEN_mainform.tpage_mat.Hide();
            _AGEN_mainform.tpage_cust_scan.Hide();
            _AGEN_mainform.tpage_cust_draw.Hide();
            _AGEN_mainform.tpage_sheet_gen.Hide();


            _AGEN_mainform.tpage_owner_scan.Show();

        }
        public string get_comboBox_prop_block()
        {
            return comboBox_prop_block.Text;
        }
        public string get_comboBox_prop_atr_sta1()
        {
            return comboBox_prop_atr_sta1.Text;
        }
        public string get_comboBox_prop_atr_sta2()
        {
            return comboBox_prop_atr_sta2.Text;
        }
        public string get_comboBox_prop_atr_distance()
        {
            return comboBox_prop_atr_distance.Text;
        }
        public string get_comboBox_prop_atr_linelist()
        {
            return comboBox_prop_atr_linelist.Text;
        }
        public string get_comboBox_prop_atr_owner()
        {
            return comboBox_prop_atr_owner.Text;
        }


        public void set_comboBox_prop_block(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_block.Items.Contains(txt) == false)
                {
                    comboBox_prop_block.Items.Add(txt);
                }
                comboBox_prop_block.SelectedIndex = comboBox_prop_block.Items.IndexOf(txt);
            }
        }
        public void set_comboBox_prop_atr_sta1(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_atr_sta1.Items.Contains(txt) == false)
                {
                    comboBox_prop_atr_sta1.Items.Add(txt);
                }
                comboBox_prop_atr_sta1.SelectedIndex = comboBox_prop_atr_sta1.Items.IndexOf(txt);
            }
        }
        public void set_comboBox_prop_atr_sta2(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_atr_sta2.Items.Contains(txt) == false)
                {
                    comboBox_prop_atr_sta2.Items.Add(txt);
                }
                comboBox_prop_atr_sta2.SelectedIndex = comboBox_prop_atr_sta2.Items.IndexOf(txt);
            }
        }
        public void set_comboBox_prop_atr_distance(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_atr_distance.Items.Contains(txt) == false)
                {
                    comboBox_prop_atr_distance.Items.Add(txt);
                }
                comboBox_prop_atr_distance.SelectedIndex = comboBox_prop_atr_distance.Items.IndexOf(txt);
            }
        }
        public void set_comboBox_prop_atr_linelist(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_atr_linelist.Items.Contains(txt) == false)
                {
                    comboBox_prop_atr_linelist.Items.Add(txt);
                }
                comboBox_prop_atr_linelist.SelectedIndex = comboBox_prop_atr_linelist.Items.IndexOf(txt);
            }
        }
        public void set_comboBox_prop_atr_owner(string txt)
        {
            if (txt != "")
            {
                if (comboBox_prop_atr_owner.Items.Contains(txt) == false)
                {
                    comboBox_prop_atr_owner.Items.Add(txt);
                }
                comboBox_prop_atr_owner.SelectedIndex = comboBox_prop_atr_owner.Items.IndexOf(txt);
            }
        }

        private Point3d get_band_insertion_point_and_band_height()
        {
            Point3d pt_ins = new Point3d();

            string band_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(3);


            if (_AGEN_mainform.Data_Table_regular_bands != null)
            {
                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                {
                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                    {
                        if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                        {
                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                            if (bn != null)
                            {
                                if (bn == band_name)
                                {
                                    if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"] != DBNull.Value &&
                                        _AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"] != DBNull.Value)
                                    {

                                        double bandh = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_height"]);
                                        double x0 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_x"]);
                                        double y0 = Convert.ToDouble(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["viewport_ms_y"]);

                                        pt_ins = new Point3d(x0, y0, bandh);

                                        if (comboBox_prop_block.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_name"] = comboBox_prop_block.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_name"] = DBNull.Value;
                                        }

                                        if (comboBox_prop_atr_sta1.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr1"] = comboBox_prop_atr_sta1.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr1"] = DBNull.Value;
                                        }


                                        if (comboBox_prop_atr_sta2.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr2"] = comboBox_prop_atr_sta2.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr2"] = DBNull.Value;
                                        }


                                        if (comboBox_prop_atr_distance.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_len_atr"] = comboBox_prop_atr_distance.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_len_atr"] = DBNull.Value;
                                        }


                                        if (comboBox_prop_atr_linelist.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field1"] = comboBox_prop_atr_linelist.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field1"] = DBNull.Value;
                                        }


                                        if (comboBox_prop_atr_owner.Text != "")
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field2"] = comboBox_prop_atr_owner.Text;
                                        }
                                        else
                                        {
                                            _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field2"] = DBNull.Value;
                                        }


                                    }
                                }
                            }
                        }
                    }
                }
            }


            return pt_ins;
        }


        private void button_insert_prop_band_Click(object sender, EventArgs e)
        {
            int index_err = 0;

            


            _AGEN_mainform Ag = this.MdiParent as _AGEN_mainform;
            _AGEN_mainform.tpage_processing.Show();
            // Ag.WindowState = FormWindowState.Minimized;

            if (get_comboBox_prop_block() == "")
            {
                _AGEN_mainform.tpage_processing.Hide();
                MessageBox.Show("you did not specified property block name");
                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }

            int lr = 1;
            if (_AGEN_mainform.Left_to_Right == false) lr = -1;

            _AGEN_mainform.Point0_prop = get_band_insertion_point_and_band_height();
            double prop_band_height = _AGEN_mainform.Point0_prop.Z;

            _AGEN_mainform.Point0_prop = new Point3d(_AGEN_mainform.Point0_prop.X, _AGEN_mainform.Point0_prop.Y, 0);

            string ProjF = _AGEN_mainform.tpage_setup.Get_project_database_folder();
            if (ProjF.Substring(ProjF.Length - 1, 1) != "\\")
            {
                ProjF = ProjF + "\\";
            }

            string fisier_prop = ProjF + _AGEN_mainform.property_excel_name;

            if (System.IO.File.Exists(fisier_prop) == false)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the property data file does not exist");
                return;
            }

            string fisier_si = ProjF + _AGEN_mainform.sheet_index_excel_name;

            if (System.IO.File.Exists(fisier_si) == false)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the sheet index data file does not exist");
                return;
            }

            string fisier_cl = ProjF + _AGEN_mainform.cl_excel_name;

            if (System.IO.File.Exists(fisier_cl) == false)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the centerline data file does not exist");
                _AGEN_mainform.dt_station_equation = null;
                return;
            }


            if (prop_band_height <= 0)
            {
                MessageBox.Show("you did not picked the ownership band");
                _AGEN_mainform.tpage_processing.Hide();
                _AGEN_mainform.tpage_blank.Hide();
                _AGEN_mainform.tpage_setup.Hide();

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


                _AGEN_mainform.tpage_viewport_settings.Show();


                Ag.WindowState = FormWindowState.Normal;
                set_enable_true();
                return;
            }


            set_enable_false(sender);

            if (System.IO.Directory.Exists(_AGEN_mainform.tpage_setup.Get_project_database_folder()) == true)
            {


                string prop_band_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(3);

                int index_property = -1;

                bool band_found = false;



                if (prop_band_name != "")
                {
                    if (_AGEN_mainform.Data_Table_regular_bands != null)
                    {
                        if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                        {

                            for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                            {
                                if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                {
                                    string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                                    if (bn == prop_band_name)
                                    {
                                        index_property = i;
                                        band_found = true;

                                        i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                    }


                                }
                            }
                        }
                    }
                }


                if (band_found == false)
                {
                    MessageBox.Show("the property band it is not defined");
                    _AGEN_mainform.tpage_processing.Hide();
                    set_enable_true();
                    return;
                }

                _AGEN_mainform.Data_Table_property = _AGEN_mainform.tpage_setup.Load_existing_property(fisier_prop);

                if (_AGEN_mainform.dt_sheet_index == null)
                {
                    _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);
                }
                else
                {
                    if (_AGEN_mainform.dt_sheet_index.Rows.Count == 0)
                    {
                        _AGEN_mainform.dt_sheet_index = _AGEN_mainform.tpage_setup.Load_existing_sheet_index(fisier_si);
                    }
                }



                if (_AGEN_mainform.dt_centerline == null || _AGEN_mainform.dt_centerline.Rows.Count == 0)
                {
                    _AGEN_mainform.tpage_setup.Load_centerline_and_station_equation(fisier_cl);
                }


            }
            else
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the project folder does not exist");
                return;
            }

            if (_AGEN_mainform.Data_Table_property.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the property file does not have any data");
                return;
            }



            if (_AGEN_mainform.dt_centerline.Rows.Count == 0)
            {
                _AGEN_mainform.tpage_processing.Hide();
                set_enable_true();
                MessageBox.Show("the centerline file does not have any data");
                return;
            }

            Functions.Load_entities_records_from_config_file(_AGEN_mainform.config_path);

            string Sta1 = "Sta1";
            string Sta2 = "Sta2";
            string Sta1CSF = "Sta1CSF";
            string Sta2CSF = "Sta2CSF";

            string colm1 = "M1";
            string colm2 = "M2";
            string colm1csf = "M1csf";
            string colm2csf = "M2csf";

            string ll_col = "LineList";
            string own_col = "Owner";

            string Pageno = "Page";
            string Rect_len = "RectangleML";

            string stretch_val = "StrechVal";
            string BandL = "BandL";
            string DeltaX_col = "DeltaX";
            string stretch_val_orig = "StrechValoriginal";

            System.Data.DataTable Data_table_compiled = new System.Data.DataTable();
            Data_table_compiled.Columns.Add(_AGEN_mainform.Col_dwg_name, typeof(string));
            Data_table_compiled.Columns.Add(Sta1, typeof(double));
            Data_table_compiled.Columns.Add(Sta2, typeof(double));
            Data_table_compiled.Columns.Add(ll_col, typeof(string));
            Data_table_compiled.Columns.Add(own_col, typeof(string));
            Data_table_compiled.Columns.Add(Pageno, typeof(int));
            Data_table_compiled.Columns.Add(Rect_len, typeof(double));
            Data_table_compiled.Columns.Add(BandL, typeof(double));
            Data_table_compiled.Columns.Add(DeltaX_col, typeof(double));
            Data_table_compiled.Columns.Add(colm1, typeof(double));
            Data_table_compiled.Columns.Add(colm2, typeof(double));
            Data_table_compiled.Columns.Add(stretch_val, typeof(double));
            Data_table_compiled.Columns.Add(stretch_val_orig, typeof(double));
            Data_table_compiled.Columns.Add("dt_row", typeof(int));
            Data_table_compiled.Columns.Add("xbeg", typeof(double));
            Data_table_compiled.Columns.Add("ybeg", typeof(double));
            Data_table_compiled.Columns.Add("xend", typeof(double));
            Data_table_compiled.Columns.Add("yend", typeof(double));
            Data_table_compiled.Columns.Add(Sta1CSF, typeof(double));
            Data_table_compiled.Columns.Add(Sta2CSF, typeof(double));
            Data_table_compiled.Columns.Add(colm1csf, typeof(double));
            Data_table_compiled.Columns.Add(colm2csf, typeof(double));



            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        Autodesk.AutoCAD.DatabaseServices.BlockTable BlockTable_data1 = ThisDrawing.Database.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
                        Autodesk.AutoCAD.DatabaseServices.BlockTableRecord BTrecord = Trans1.GetObject(ThisDrawing.Database.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;
                        string Block_name = get_comboBox_prop_block();
                        if (BlockTable_data1.Has(Block_name) == false)
                        {
                            MessageBox.Show("the block name you specified does not belong to the current drawing");
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            return;
                        }
                        Polyline3d poly3d = null;
                        Polyline poly2d = Functions.Build_2D_CL_from_dt_cl(_AGEN_mainform.dt_centerline);
                        double poly_length = poly2d.Length;
                        if (_AGEN_mainform.Project_type == "3D")
                        {
                            poly3d = Functions.Build_3d_poly_for_scanning(_AGEN_mainform.dt_centerline);
                            poly_length = poly3d.Length;

                        }


                        if (_AGEN_mainform.dt_sheet_index == null || _AGEN_mainform.dt_sheet_index.Rows.Count == 0)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            MessageBox.Show("the data of sheet index file is not complete");
                            return;
                        }

                        #region USA
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
                                        double eq_meas = poly2d.GetDistAtPoint(pt_on_2d);
                                        if (_AGEN_mainform.Project_type == "3D")
                                        {
                                            double param1 = poly2d.GetParameterAtPoint(pt_on_2d);
                                            eq_meas = poly3d.GetDistanceAtParameter(param1);
                                        }

                                        _AGEN_mainform.dt_station_equation.Rows[i]["measured"] = eq_meas;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_AGEN_mainform.COUNTRY == "USA") _AGEN_mainform.dt_station_equation = null;
                        }
                        #endregion



                        List<int> lista_bands_for_generation = new List<int>();

                        if (comboBox_start.Text == "" || comboBox_end.Text == "")
                        {
                            lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg("", "");
                        }


                        if (comboBox_start.Text != "" & comboBox_end.Text != "")
                        {
                            lista_bands_for_generation = _AGEN_mainform.tpage_setup.create_band_list_of_dwg(comboBox_start.Text, comboBox_end.Text);
                        }

                        if (lista_bands_for_generation.Count == 0)
                        {
                            _AGEN_mainform.tpage_processing.Hide();
                            set_enable_true();
                            MessageBox.Show("please check your input");
                            return;
                        }



                        double Min_dist = 0;
                        BlockReference BR1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, new Point3d(0, 0, 0), 1, 0, "0", new System.Collections.Specialized.StringCollection(), new System.Collections.Specialized.StringCollection());
                        Min_dist = Functions.Get_distance1_block(BR1);
                        BR1.Erase();

                        Functions.Creaza_layer(_AGEN_mainform.layer_ownership_band_no_plot, 30, false);



                        string prop_band_name = _AGEN_mainform.tpage_viewport_settings.get_comboBox_viewport_target_areas(3);


                        if (prop_band_name != "")
                        {
                            if (_AGEN_mainform.Data_Table_regular_bands != null)
                            {
                                if (_AGEN_mainform.Data_Table_regular_bands.Rows.Count > 0)
                                {

                                    for (int i = 0; i < _AGEN_mainform.Data_Table_regular_bands.Rows.Count; ++i)
                                    {
                                        if (_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"] != DBNull.Value)
                                        {
                                            string bn = Convert.ToString(_AGEN_mainform.Data_Table_regular_bands.Rows[i]["band_name"]);
                                            if (bn == prop_band_name)
                                            {

                                                if (comboBox_prop_block.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_name"] = comboBox_prop_block.Text;
                                                }

                                                if (comboBox_prop_atr_sta1.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr1"] = comboBox_prop_atr_sta1.Text;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr1"] = DBNull.Value;
                                                }

                                                if (comboBox_prop_atr_sta2.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr2"] = comboBox_prop_atr_sta2.Text;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_sta_atr2"] = DBNull.Value;
                                                }

                                                if (comboBox_prop_atr_distance.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_len_atr"] = comboBox_prop_atr_distance.Text;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_len_atr"] = DBNull.Value;
                                                }

                                                if (comboBox_prop_atr_linelist.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field1"] = comboBox_prop_atr_linelist.Text;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field1"] = DBNull.Value;
                                                }

                                                if (comboBox_prop_atr_owner.Text != "")
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field2"] = comboBox_prop_atr_owner.Text;
                                                }
                                                else
                                                {
                                                    _AGEN_mainform.Data_Table_regular_bands.Rows[i]["block_field2"] = DBNull.Value;
                                                }


                                                i = _AGEN_mainform.Data_Table_regular_bands.Rows.Count;
                                            }


                                        }
                                    }
                                }
                            }
                        }


                        for (int i = 0; i < _AGEN_mainform.Data_Table_property.Rows.Count; ++i)
                        {
                            _AGEN_mainform.Data_Table_property.Rows[i]["BlockHandle"] = DBNull.Value;
                            int m_start = 0;
                            bool Boolean_go_to_check_s1_s2 = false;
                            double Station1 = -1.123;
                            double Station2 = -1.123;
                            double Station1_CSF = -1.123;
                            double Station2_CSF = -1.123;

                            double px_start = -1.2345;
                            double py_start = -1.2345;
                            double px_end = -1.2345;
                            double py_end = -1.2345;





                            if (_AGEN_mainform.Project_type == "2D")
                            {
                                if (_AGEN_mainform.Data_Table_property.Rows[i]["2DStaBeg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["2DStaEnd"] != DBNull.Value)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["2DStaBeg"]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["2DStaEnd"]);


                                    if (poly_length < Station1) Station1 = poly_length;
                                    if (poly_length < Station2) Station2 = poly_length;
                                    if (Station1 < 0) Station1 = 0;
                                    if (Station2 < 0) Station2 = 0;

                                    if (_AGEN_mainform.COUNTRY == "USA")
                                    {
                                        Station1_CSF = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                        Station2_CSF = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);
                                    }
                                    else
                                    {
                                        Station1_CSF = Station1;
                                        Station2_CSF = Station2;
                                    }
                                }
                            }
                            else
                            {
                                if (_AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"] != DBNull.Value && _AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"] != DBNull.Value &&
                                    Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"])) == true && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"])) == true)
                                {
                                    Station1 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["3DStaBeg"]);
                                    Station2 = Convert.ToDouble(_AGEN_mainform.Data_Table_property.Rows[i]["3DStaEnd"]);

                                    if (poly_length < Station1) Station1 = poly_length;
                                    if (poly_length < Station2) Station2 = poly_length;
                                    if (Station1 < 0) Station1 = 0;
                                    if (Station2 < 0) Station2 = 0;
                                    if (_AGEN_mainform.COUNTRY == "USA")
                                    {
                                        Station1_CSF = Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation);
                                        Station2_CSF = Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation);
                                    }
                                    else
                                    {
                                        Station1_CSF = Station1;
                                        Station2_CSF = Station2;
                                    }
                                }
                            }



                            if (Station1 != -1.123 && Station2 != -1.123)
                            {

                                string Parcelid = "NoData";
                                string Owner1 = "NoData";
                                if (_AGEN_mainform.Data_Table_property.Rows[i][_AGEN_mainform.Col_Owner] != DBNull.Value)
                                {
                                    Owner1 = _AGEN_mainform.Data_Table_property.Rows[i][_AGEN_mainform.Col_Owner].ToString();
                                }
                                if (_AGEN_mainform.Data_Table_property.Rows[i][_AGEN_mainform.Col_Linelist] != DBNull.Value)
                                {
                                    Parcelid = _AGEN_mainform.Data_Table_property.Rows[i][_AGEN_mainform.Col_Linelist].ToString();
                                }

                                double mx_start = -1.2345;
                                double my_start = -1.2345;
                                double mx_end = -1.2345;
                                double my_end = -1.2345;

                            L123:

                                for (int j = m_start; j < _AGEN_mainform.dt_sheet_index.Rows.Count; ++j)
                                {
                                    if (_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2] != DBNull.Value
                                        && _AGEN_mainform.dt_sheet_index.Rows[j]["X_Beg"] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j]["Y_Beg"] != DBNull.Value
                                        && _AGEN_mainform.dt_sheet_index.Rows[j]["X_End"] != DBNull.Value && _AGEN_mainform.dt_sheet_index.Rows[j]["Y_End"] != DBNull.Value
                                        && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1])) == true
                                         && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2])) == true
                                          && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["X_Beg"])) == true
                                           && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_Beg"])) == true
                                            && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["X_End"])) == true
                                             && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_End"])) == true)
                                    {
                                        double M1 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M1]);
                                        double M2 = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_M2]);

                                        double M1_CSF = M1;
                                        double M2_CSF = M2;

                                        if (_AGEN_mainform.dt_sheet_index.Columns.Contains("M1_CANADA") &&
                                            _AGEN_mainform.dt_sheet_index.Rows[j]["M1_CANADA"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["M1_CANADA"])) == true)
                                        {
                                            M1_CSF = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M1_CANADA"]);
                                        }

                                        if (_AGEN_mainform.dt_sheet_index.Columns.Contains("M2_CANADA") &&
                                            _AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"] != DBNull.Value &&
                                            Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"])) == true)
                                        {
                                            M2_CSF = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["M2_CANADA"]);
                                        }

                                        if (_AGEN_mainform.COUNTRY == "USA")
                                        {
                                            M1_CSF = Functions.Station_equation_ofV2(M1, _AGEN_mainform.dt_station_equation);
                                            M2_CSF = Functions.Station_equation_ofV2(M2, _AGEN_mainform.dt_station_equation);
                                        }

                                        mx_start = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["X_Beg"]);
                                        my_start = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_Beg"]);
                                        mx_end = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["X_End"]);
                                        my_end = Convert.ToDouble(_AGEN_mainform.dt_sheet_index.Rows[j]["Y_End"]);

                                        if (M2 <= M1)
                                        {
                                            _AGEN_mainform.tpage_processing.Hide();
                                            set_enable_true();
                                            MessageBox.Show("End Station is smaller than Start Station on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                            return;
                                        }

                                        if (M2 > poly_length)
                                        {
                                            if (Math.Abs(M2 - poly_length) < 0.99)
                                            {
                                                M2 = poly_length;
                                            }
                                            else
                                            {
                                                _AGEN_mainform.tpage_processing.Hide();
                                                set_enable_true();
                                                MessageBox.Show("End Station is bigger than poly length on row " + (j).ToString() + "\r\n" + _AGEN_mainform.sheet_index_excel_name);
                                                return;
                                            }
                                        }

                                        Point3d pm1 = new Point3d();
                                        Point3d pm2 = new Point3d();

                                        if (M1 > poly_length) M1 = poly_length;
                                        if (M2 > poly_length) M2 = poly_length;

                                        try
                                        {
                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                pm1 = poly3d.GetPointAtDist(M1);
                                            }
                                            else
                                            {
                                                pm1 = poly2d.GetPointAtDist(M1);
                                            }

                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                        {
                                            pm1 = poly2d.EndPoint;
                                        }

                                        try
                                        {

                                            if (_AGEN_mainform.Project_type == "3D")
                                            {
                                                pm2 = poly3d.GetPointAtDist(M2);
                                            }
                                            else
                                            {
                                                pm2 = poly2d.GetPointAtDist(M2);
                                            }
                                        }
                                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                        {
                                            pm2 = poly2d.EndPoint;
                                        }


                                        pm1 = new Point3d(pm1.X, pm1.Y, 0);
                                        pm2 = new Point3d(pm2.X, pm2.Y, 0);

                                        Line Linie_M1_M2 = new Line(new Point3d(pm1.X, pm1.Y, 0), new Point3d(pm2.X, pm2.Y, 0));

                                        if (Boolean_go_to_check_s1_s2 == true)
                                        {
                                            if (Math.Round(Station1, 0) == Math.Round(Station2, 0))
                                            {
                                                goto LS12end;
                                            }
                                            goto LS1S2;
                                        }


                                        if (Math.Round(M1, 4) <= Math.Round(Station1, 4) && Math.Round(M2, 4) <= Math.Round(Station2, 4) && Math.Round(M1, 4) <= Math.Round(Station2, 4) && Math.Round(M2, 4) > Math.Round(Station1, 4))
                                        {

                                            Point3d ppt1 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.GetPointAtDist(Station1);
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.GetPointAtDist(Station1);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt1 = poly2d.EndPoint;
                                            }
                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][ll_col] = Parcelid;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][own_col] = Owner1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = mx_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = my_end;

                                            Station1 = M2;
                                            Station1_CSF = M2_CSF;

                                            px_start = mx_end;
                                            py_start = my_end;

                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }

                                        if (Math.Round(Station1, 4) >= Math.Round(M1, 4) && Math.Round(Station2, 4) <= Math.Round(M2, 4) && Math.Round(Station1, 4) < Math.Round(M2, 4))
                                        {

                                            Point3d ppt1 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.GetPointAtDist(Station1);
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.GetPointAtDist(Station1);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt1 = poly2d.EndPoint;
                                            }


                                            Point3d ppt2 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt2 = poly3d.GetPointAtDist(Station2);
                                                }
                                                else
                                                {
                                                    ppt2 = poly2d.GetPointAtDist(Station2);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt2 = poly2d.EndPoint;
                                            }

                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(ppt2, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = Station2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = Station2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][ll_col] = Parcelid;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][own_col] = Owner1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = px_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = py_end;

                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }
                                    LS1S2:
                                        if (Math.Round(Station1, 4) >= Math.Round(M1, 4) && Math.Round(Station2, 4) <= Math.Round(M2, 4) && Math.Round(Station1, 4) < Math.Round(M2, 4))
                                        {
                                            Point3d ppt1 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.GetPointAtDist(Station1);
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.GetPointAtDist(Station1);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt1 = poly2d.EndPoint;
                                            }


                                            Point3d ppt2 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt2 = poly3d.GetPointAtDist(Station2);
                                                }
                                                else
                                                {
                                                    ppt2 = poly2d.GetPointAtDist(Station2);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt2 = poly2d.EndPoint;
                                            }


                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            Point3d Pt2 = Linie_M1_M2.GetClosestPointTo(ppt2, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(Pt2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = Station2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = Station2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][ll_col] = Parcelid;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][own_col] = Owner1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = px_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = py_end;

                                            j = _AGEN_mainform.dt_sheet_index.Rows.Count;
                                            goto LS12end;
                                        }
                                        else if (Math.Round(Station1, 4) < Math.Round(M2, 4) && Math.Round(Station1, 4) >= Math.Round(M1, 4))
                                        {
                                            Point3d ppt1 = new Point3d();
                                            try
                                            {
                                                if (_AGEN_mainform.Project_type == "3D")
                                                {
                                                    ppt1 = poly3d.GetPointAtDist(Station1);
                                                }
                                                else
                                                {
                                                    ppt1 = poly2d.GetPointAtDist(Station1);
                                                }

                                            }
                                            catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                            {
                                                ppt1 = poly2d.EndPoint;
                                            }

                                            Point3d Pt1 = Linie_M1_M2.GetClosestPointTo(ppt1, Vector3d.ZAxis, false);
                                            double stretch01 = Pt1.DistanceTo(pm2) * _AGEN_mainform.Vw_scale;
                                            double deltax1 = Pt1.DistanceTo(pm1) * _AGEN_mainform.Vw_scale;

                                            Data_table_compiled.Rows.Add();
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][_AGEN_mainform.Col_dwg_name] = _AGEN_mainform.dt_sheet_index.Rows[j][_AGEN_mainform.Col_dwg_name];
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1] = Station1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta1CSF] = Station1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Sta2CSF] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][ll_col] = Parcelid;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][own_col] = Owner1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Pageno] = j + 1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1] = M1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2] = M2;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm1csf] = M1_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][colm2csf] = M2_CSF;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][Rect_len] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][stretch_val_orig] = stretch01;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][BandL] = Linie_M1_M2.Length * _AGEN_mainform.Vw_scale;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1][DeltaX_col] = deltax1;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["dt_row"] = i;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xbeg"] = px_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["ybeg"] = py_start;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["xend"] = mx_end;
                                            Data_table_compiled.Rows[Data_table_compiled.Rows.Count - 1]["yend"] = my_end;

                                            Station1 = M2;
                                            Station1_CSF = M2_CSF;
                                            px_start = mx_end;
                                            py_start = my_end;



                                            m_start = j + 1;
                                            Boolean_go_to_check_s1_s2 = true;
                                            goto L123;
                                        }
                                    }
                                }
                            LS12end:
                                string xx = "";
                            }
                        }



                        // Alignment_generator.Functions.Transfer_datatable_to_new_excel_spreadsheet(Data_table_compiled);

                        Functions.Creaza_layer(_AGEN_mainform.layer_ownership_band_no_plot, 30, false);


                        Functions.Creaza_layer(_AGEN_mainform.layer_ownership_band, 7, true);


                        int Pagep = -1;

                        if (Data_table_compiled != null)
                        {
                            Autodesk.Gis.Map.ObjectData.Tables Tables1 = Autodesk.Gis.Map.HostMapApplicationServices.Application.ActiveProject.ODTables;
                            Functions.Create_ownership_od_table();

                            for (int i = 0; i < Data_table_compiled.Rows.Count; ++i)
                            {
                                int Page1 = Convert.ToInt32(Data_table_compiled.Rows[i][Pageno]);
                                double ml_len = Convert.ToDouble(Data_table_compiled.Rows[i][Rect_len]);
                                string dwg_name = Data_table_compiled.Rows[i][_AGEN_mainform.Col_dwg_name].ToString();
                                double strech1 = Convert.ToDouble(Data_table_compiled.Rows[i][stretch_val]);
                                double Diff = Min_dist - strech1;
                                if (Diff > 0)
                                {
                                    Data_table_compiled.Rows[i][stretch_val] = Min_dist;

                                    for (int j = 0; j < Data_table_compiled.Rows.Count; ++j)
                                    {
                                        int Page2 = Convert.ToInt32(Data_table_compiled.Rows[j][Pageno]);
                                        double deltax2 = Convert.ToDouble(Data_table_compiled.Rows[j][DeltaX_col]);
                                        double band_len2 = Convert.ToDouble(Data_table_compiled.Rows[j][BandL]);

                                        if (Page1 == Page2)
                                        {
                                            Data_table_compiled.Rows[j][BandL] = band_len2 + Diff;

                                            if (i < j)
                                            {
                                                Data_table_compiled.Rows[j][DeltaX_col] = deltax2 + Diff;
                                            }
                                        }
                                    }
                                }
                                if (Page1 != Pagep)
                                {
                                    if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                    {
                                        Polyline vp_vw1 = new Polyline();

                                        vp_vw1.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_prop.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_prop.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_prop.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_prop.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_prop.X + _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_prop.Y - prop_band_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw1.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_prop.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_prop.Y - prop_band_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                        vp_vw1.Closed = true;
                                        vp_vw1.Layer = _AGEN_mainform.layer_ownership_band_no_plot;
                                        vp_vw1.ColorIndex = 3;
                                        BTrecord.AppendEntity(vp_vw1);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw1, true);

                                        Polyline vp_vw2 = new Polyline();

                                        vp_vw2.AddVertexAt(0, new Point2d(_AGEN_mainform.Point0_prop.X - ml_len / 2, _AGEN_mainform.Point0_prop.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(1, new Point2d(_AGEN_mainform.Point0_prop.X + ml_len / 2, _AGEN_mainform.Point0_prop.Y - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(2, new Point2d(_AGEN_mainform.Point0_prop.X + ml_len / 2, _AGEN_mainform.Point0_prop.Y - prop_band_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);
                                        vp_vw2.AddVertexAt(3, new Point2d(_AGEN_mainform.Point0_prop.X - ml_len / 2, _AGEN_mainform.Point0_prop.Y - prop_band_height - (Page1 - 1) * _AGEN_mainform.Band_Separation), 0, 0, 0);

                                        vp_vw2.Closed = true;
                                        vp_vw2.Layer = _AGEN_mainform.layer_ownership_band_no_plot;
                                        vp_vw2.ColorIndex = 1;
                                        BTrecord.AppendEntity(vp_vw2);
                                        Trans1.AddNewlyCreatedDBObject(vp_vw2, true);

                                        MText Band_label = new MText();
                                        Band_label.Contents = dwg_name;
                                        Band_label.TextHeight = prop_band_height / 2;
                                        Band_label.Rotation = 0;
                                        Band_label.Attachment = AttachmentPoint.MiddleLeft;
                                        Band_label.Location = new Point3d(_AGEN_mainform.Point0_prop.X - _AGEN_mainform.Vw_width / 2, _AGEN_mainform.Point0_prop.Y - prop_band_height / 2 - (Page1 - 1) * _AGEN_mainform.Band_Separation, 0);
                                        Band_label.Layer = _AGEN_mainform.layer_ownership_band_no_plot;

                                        double gap1 = (_AGEN_mainform.Vw_width - ml_len * _AGEN_mainform.Vw_scale) / 2;
                                        Extents3d gerect = Band_label.GeometricExtents;
                                        Point3d p2 = gerect.MaxPoint;
                                        Point3d p1 = gerect.MinPoint;
                                        bool repeat1 = false;
                                        do
                                        {
                                            if (p2.X - p1.X > gap1 - prop_band_height / 3 && Band_label.TextHeight >= 2)
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


                                    }


                                    Pagep = Page1;
                                }
                            }

                            #region data table compiled

                            double xp = -1.2345;
                            double stretchp = -1.2345;
                            int pagep = -2;
                            string visib1 = "";



                            for (int i = 0; i < Data_table_compiled.Rows.Count; ++i)
                            {
                                index_err = i;
                                int Page1 = Convert.ToInt32(Data_table_compiled.Rows[i][Pageno]);

                                if (lista_bands_for_generation.Contains(Page1 - 1) == true)
                                {
                                    double Station1 = Math.Round(Convert.ToDouble(Data_table_compiled.Rows[i][Sta1]), _AGEN_mainform.round1);
                                    double Station2 = Math.Round(Convert.ToDouble(Data_table_compiled.Rows[i][Sta2]), _AGEN_mainform.round1);
                                    if (Station1 > poly_length) Station1 = poly_length;
                                    if (Station2 >= poly_length) Station2 = poly_length;

                                    double M1 = Convert.ToDouble(Data_table_compiled.Rows[i][colm1]);
                                    double M2 = Convert.ToDouble(Data_table_compiled.Rows[i][colm2]);
                                    double Station1_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][Sta1CSF]);
                                    double Station2_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][Sta2CSF]);
                                    double M1_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][colm1csf]);
                                    double M2_CSF = Convert.ToDouble(Data_table_compiled.Rows[i][colm2csf]);

                                    string Owner1 = Data_table_compiled.Rows[i][own_col].ToString();
                                    string Linelist = Data_table_compiled.Rows[i][ll_col].ToString();

                                    double ml_len = Convert.ToDouble(Data_table_compiled.Rows[i][Rect_len]);
                                    double band_len = Convert.ToDouble(Data_table_compiled.Rows[i][BandL]);
                                    double Diff = (band_len - ml_len) / 2;
                                    double deltax = Convert.ToDouble(Data_table_compiled.Rows[i][DeltaX_col]);

                                    string sta1_string = "-1";
                                    string sta2_string = "-1";

                                    if (_AGEN_mainform.COUNTRY == "USA")
                                    {
                                        sta1_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Station1, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                        sta2_string = Functions.Get_chainage_from_double(Functions.Station_equation_ofV2(Station2, _AGEN_mainform.dt_station_equation), _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);

                                    }
                                    else
                                    {
                                        sta1_string = Functions.Get_chainage_from_double(Station1_CSF, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);
                                        sta2_string = Functions.Get_chainage_from_double(Station2_CSF, _AGEN_mainform.units_of_measurement, _AGEN_mainform.round1);


                                    }


                                    string Suff = "'";
                                    if (_AGEN_mainform.units_of_measurement == "m") Suff = "";

                                    string len1 = Functions.Get_String_Rounded((Math.Round(Station2, _AGEN_mainform.round1) - Math.Round(Station1, _AGEN_mainform.round1)), _AGEN_mainform.round1) + Suff;

                                    double strech1 = Convert.ToDouble(Data_table_compiled.Rows[i][stretch_val]);

                                    double x = _AGEN_mainform.Point0_prop.X - lr * ml_len / 2 + lr * deltax;
                                    double y = _AGEN_mainform.Point0_prop.Y - prop_band_height - (Page1 - 1) * _AGEN_mainform.Band_Separation;

                                    double x_beg = -1.2345;
                                    double y_beg = -1.2345;
                                    double x_end = -1.2345;
                                    double y_end = -1.2345;

                                    if (Data_table_compiled.Rows[i]["xbeg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Data_table_compiled.Rows[i]["xbeg"])) == true)
                                    {
                                        x_beg = Convert.ToDouble(Data_table_compiled.Rows[i]["xbeg"]);
                                    }

                                    if (Data_table_compiled.Rows[i]["ybeg"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Data_table_compiled.Rows[i]["ybeg"])) == true)
                                    {
                                        y_beg = Convert.ToDouble(Data_table_compiled.Rows[i]["ybeg"]);
                                    }

                                    if (Data_table_compiled.Rows[i]["xend"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Data_table_compiled.Rows[i]["xend"])) == true)
                                    {
                                        x_end = Convert.ToDouble(Data_table_compiled.Rows[i]["xend"]);
                                    }

                                    if (Data_table_compiled.Rows[i]["yend"] != DBNull.Value && Functions.IsNumeric(Convert.ToString(Data_table_compiled.Rows[i]["yend"])) == true)
                                    {
                                        y_end = Convert.ToDouble(Data_table_compiled.Rows[i]["yend"]);
                                    }

                                    if (_AGEN_mainform.dt_config_ownership != null && _AGEN_mainform.dt_config_ownership.Rows.Count > 0)
                                    {
                                        for (int j = 0; j < _AGEN_mainform.dt_config_ownership.Rows.Count; ++j)
                                        {
                                            double cx_beg = -1.2345;
                                            double cy_beg = -1.2345;
                                            double cx_end = -1.2345;
                                            double cy_end = -1.2345;

                                            if (_AGEN_mainform.dt_config_ownership.Rows[j][12] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][12])) == true)
                                            {
                                                cx_beg = Convert.ToDouble(_AGEN_mainform.dt_config_ownership.Rows[j][12]);
                                            }

                                            if (_AGEN_mainform.dt_config_ownership.Rows[j][13] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][13])) == true)
                                            {
                                                cy_beg = Convert.ToDouble(_AGEN_mainform.dt_config_ownership.Rows[j][13]);
                                            }

                                            if (_AGEN_mainform.dt_config_ownership.Rows[j][14] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][14])) == true)
                                            {
                                                cx_end = Convert.ToDouble(_AGEN_mainform.dt_config_ownership.Rows[j][14]);
                                            }

                                            if (_AGEN_mainform.dt_config_ownership.Rows[j][15] != DBNull.Value && Functions.IsNumeric(Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][15])) == true)
                                            {
                                                cy_end = Convert.ToDouble(_AGEN_mainform.dt_config_ownership.Rows[j][15]);
                                            }

                                            if (cx_beg != -1.2345 && cy_beg != -1.2345 && cx_end != -1.2345 && cy_end != -1.2345 && x_beg != -1.2345 && y_beg != -1.2345 && x_end != -1.2345 && y_end != -1.2345)
                                            {
                                                if (Math.Abs(cx_beg - x_beg) < 0.1 && Math.Abs(cy_beg - y_beg) < 0.1 && Math.Abs(cx_end - x_end) < 0.1 && Math.Abs(cy_end - y_end) < 0.1)
                                                {
                                                    if (_AGEN_mainform.dt_config_ownership.Rows[j][6] != DBNull.Value)
                                                    {
                                                        strech1 = Convert.ToDouble(_AGEN_mainform.dt_config_ownership.Rows[j][6]);
                                                        if (xp != -1.2345 && stretchp != -1.2345 && pagep == Page1)
                                                        {
                                                            x = xp + stretchp;
                                                        }
                                                    }
                                                    if (_AGEN_mainform.dt_config_ownership.Rows[j][5] != DBNull.Value)
                                                    {
                                                        visib1 = Convert.ToString(_AGEN_mainform.dt_config_ownership.Rows[j][5]);
                                                    }
                                                    j = _AGEN_mainform.dt_config_ownership.Rows.Count;
                                                }
                                            }
                                        }
                                    }

                                    Point3d InsPt = new Point3d(x, y, 0);

                                    System.Collections.Specialized.StringCollection Colectie_nume_atribute = new System.Collections.Specialized.StringCollection();
                                    System.Collections.Specialized.StringCollection Colectie_valori = new System.Collections.Specialized.StringCollection();

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta1());
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta2());
                                    Colectie_valori.Add(sta2_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta1() + "1");
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta2() + "1");
                                    Colectie_valori.Add(sta2_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta1() + "11");
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta2() + "11");
                                    Colectie_valori.Add(sta2_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta1() + "111");
                                    Colectie_valori.Add(sta1_string);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_sta2() + "111");
                                    Colectie_valori.Add(sta2_string);


                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_owner());
                                    Colectie_valori.Add(Owner1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_owner() + "1");
                                    Colectie_valori.Add(Owner1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_owner() + "11");
                                    Colectie_valori.Add(Owner1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_owner() + "111");
                                    Colectie_valori.Add(Owner1);


                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_linelist());
                                    Colectie_valori.Add(Linelist);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_linelist() + "1");
                                    Colectie_valori.Add(Linelist);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_linelist() + "11");
                                    Colectie_valori.Add(Linelist);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_linelist() + "111");
                                    Colectie_valori.Add(Linelist);


                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_distance());
                                    Colectie_valori.Add(len1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_distance() + "1");
                                    Colectie_valori.Add(len1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_distance() + "11");
                                    Colectie_valori.Add(len1);

                                    Colectie_nume_atribute.Add(get_comboBox_prop_atr_distance() + "111");
                                    Colectie_valori.Add(len1);


                                    BlockReference Block1 = Functions.InsertBlock_with_multiple_atributes_with_database(ThisDrawing.Database, BTrecord, "", Block_name, InsPt, 1, 0,
                                                                                                                    _AGEN_mainform.layer_ownership_band, Colectie_nume_atribute, Colectie_valori);

                                    Functions.Stretch_block(Block1, "Distance1", strech1);

                                    if (visib1 != "")
                                    {
                                        Functions.set_block_visibility(Block1, visib1);
                                        visib1 = "";
                                    }

                                    List<object> Lista_val = new List<object>();
                                    List<Autodesk.Gis.Map.Constants.DataType> Lista_type = new List<Autodesk.Gis.Map.Constants.DataType>();
                                    string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                                    if (segment1 == "not defined") segment1 = "";
                                    Lista_val.Add(segment1);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);
                                    Lista_val.Add(System.DateTime.Today.Year + "-" + System.DateTime.Today.Month + "-" + System.DateTime.Today.Day + " at " + System.DateTime.Now.Hour + ":" + System.DateTime.Now.Minute);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Character);

                                    Lista_val.Add(x_beg);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add(y_beg);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add(x_end);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Lista_val.Add(y_end);
                                    Lista_type.Add(Autodesk.Gis.Map.Constants.DataType.Real);

                                    Functions.Populate_object_data_table_from_objectid(Tables1, Block1.ObjectId, "Agen_owner", Lista_val, Lista_type);

                                    xp = InsPt.X;
                                    stretchp = strech1;
                                    pagep = Convert.ToInt32(Page1);

                                    int index_prop = Convert.ToInt32(Data_table_compiled.Rows[i]["dt_row"]);
                                    string Existing_ID = "";
                                    if (_AGEN_mainform.Data_Table_property.Rows[index_prop]["BlockHandle"] != DBNull.Value)
                                    {
                                        Existing_ID = Convert.ToString(_AGEN_mainform.Data_Table_property.Rows[index_prop]["BlockHandle"]);
                                    }

                                    string New_Id = Block1.ObjectId.Handle.Value.ToString();

                                    if (Existing_ID == "")
                                    {
                                        _AGEN_mainform.Data_Table_property.Rows[index_prop]["BlockHandle"] = New_Id;
                                    }
                                    else
                                    {
                                        _AGEN_mainform.Data_Table_property.Rows[index_prop]["BlockHandle"] = Existing_ID + "," + New_Id;
                                    }



                                }
                                else
                                {
                                    xp = -1.2345;
                                    visib1 = "";
                                    stretchp = -1.2345;
                                    pagep = -2;
                                }
                            }
                        }
                        #endregion


                        write_ownership_settings_to_excel(_AGEN_mainform.config_path);
                        if (_AGEN_mainform.Project_type == "3D" && poly3d.IsErased == false) poly3d.Erase();
                        Trans1.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + index_err.ToString());
            }
            _AGEN_mainform.tpage_processing.Hide();

            set_enable_true();

            Ag.WindowState = FormWindowState.Normal;
        }

        private void button_prop_block_refresh_Click(object sender, EventArgs e)
        {
            set_enable_false(sender);
            try
            {
                Autodesk.AutoCAD.ApplicationServices.Document ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Autodesk.AutoCAD.EditorInput.Editor Editor1 = ThisDrawing.Editor;
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                using (Autodesk.AutoCAD.ApplicationServices.DocumentLock Lock1 = ThisDrawing.LockDocument())
                {
                    using (Autodesk.AutoCAD.DatabaseServices.Transaction Trans1 = ThisDrawing.TransactionManager.StartTransaction())
                    {
                        comboBox_prop_atr_distance.Items.Clear();
                        comboBox_prop_atr_linelist.Items.Clear();
                        comboBox_prop_atr_owner.Items.Clear();
                        comboBox_prop_atr_sta1.Items.Clear();
                        comboBox_prop_atr_sta2.Items.Clear();
                        Functions.Incarca_existing_Blocks_with_attributes_to_combobox(comboBox_prop_block);

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

        private void comboBox_prop_block_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Block_loaded_from_open_config == false)
            {
                Functions.Incarca_existing_Atributes_to_combobox(comboBox_prop_block.Text, comboBox_prop_atr_distance);
                Functions.Incarca_existing_Atributes_to_combobox(comboBox_prop_block.Text, comboBox_prop_atr_linelist);
                Functions.Incarca_existing_Atributes_to_combobox(comboBox_prop_block.Text, comboBox_prop_atr_owner);
                Functions.Incarca_existing_Atributes_to_combobox(comboBox_prop_block.Text, comboBox_prop_atr_sta1);
                Functions.Incarca_existing_Atributes_to_combobox(comboBox_prop_block.Text, comboBox_prop_atr_sta2);
            }
        }


        public void clear_combobox()
        {
            comboBox_prop_block.Items.Clear();
            comboBox_prop_atr_sta1.Items.Clear();
            comboBox_prop_atr_sta2.Items.Clear();
            comboBox_prop_atr_distance.Items.Clear();
            comboBox_prop_atr_linelist.Items.Clear();
            comboBox_prop_atr_owner.Items.Clear();
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
        private void label_ownership_Click(object sender, EventArgs e)
        {

        }

        public void write_ownership_settings_to_excel(string cfg1)
        {

            string ts1 = get_comboBox_prop_block();
            string ts2 = get_comboBox_prop_atr_sta1();
            string ts3 = get_comboBox_prop_atr_sta2();
            string ts4 = get_comboBox_prop_atr_distance();
            string ts5 = get_comboBox_prop_atr_linelist();
            string ts6 = get_comboBox_prop_atr_owner();


            if (ts1 != "" || ts2 != "" || ts3 != "" || ts4 != "" || ts5 != "" || ts6 != "")
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
                Microsoft.Office.Interop.Excel.Workbook Workbook1 = Excel1.Workbooks.Open(cfg1);





                Microsoft.Office.Interop.Excel.Worksheet W1 = null;


                string segment1 = _AGEN_mainform.tpage_setup.Get_segment_name1();
                if (segment1 == "not defined") segment1 = "";

                foreach (Microsoft.Office.Interop.Excel.Worksheet wsh1 in Workbook1.Worksheets)
                {
                    if (wsh1.Name == "O_dc_" + segment1)
                    {
                        W1 = wsh1;
                    }
                }

                if (W1 == null)
                {
                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W1.Name = "O_dc_" + segment1;
                }

                if (W1 == null)
                {
                    W1 = Workbook1.Worksheets.Add(System.Reflection.Missing.Value, Workbook1.Worksheets[Workbook1.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    W1.Name = "Ownership_data_config";
                }

                try
                {
                    int NrR = 9;
                    int NrC = 2;

                    Object[,] values = new object[NrR, NrC];
                    values[0, 0] = "Ownership Block Name";
                    values[0, 1] = ts1;
                    values[1, 0] = "Station start Attribute";
                    values[1, 1] = ts2;
                    values[2, 0] = "Station end Attribute";
                    values[2, 1] = ts3;
                    values[3, 0] = "Length Attribute";
                    values[3, 1] = ts4;
                    values[4, 0] = "Linelist Attribute";
                    values[4, 1] = ts5;
                    values[5, 0] = "Ownership Attribute";
                    values[5, 1] = ts6;
                    values[6, 0] = "Ownership Data Table";
                    values[7, 0] = "Owner field";
                    values[8, 0] = "Linelist(Tract) Field";


                    Microsoft.Office.Interop.Excel.Range range1 = W1.Range["A1:B9"];
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
    }
}
